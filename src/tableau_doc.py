# -*- coding: utf-8 -*-
"""
Gera documentação técnica de workbooks Tableau (.twb/.twbx).

O script:
- lê um arquivo Tableau informado pelo usuário;
- extrai todo o conteúdo do pacote para `data/<nome_do_arquivo>/`;
- gera um mapa XPath/JSON do workbook;
- gera documentação em Markdown, JSON e/ou Excel;
- usa o mesmo nome-base do arquivo de entrada nos artefatos de saída.

Opções de execução
python3 Tableau_doc.py /caminho/arquivo.twbx --format all
python3 Tableau_doc.py /caminho/arquivo.twb --format markdown
python3 Tableau_doc.py /caminho/arquivo.twbx --format json
python3 Tableau_doc.py /caminho/arquivo.twbx --format excel
Versão 1.0 - 23/03/2026
"""

from __future__ import annotations

import argparse
import base64
import json
import re
import shutil
import sys
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
DEFAULT_OUTPUT_ROOT = PROJECT_ROOT / "data"
TEMPORARY_OUTPUT_NAMES = {
    "package_contents",
    ".tmp",
    "tmp",
    "temp",
}


def sanitize_filename(value: str) -> str:
    """Converte um texto em nome de arquivo seguro, preservando legibilidade."""
    safe = re.sub(r"[^\w\-. ]+", "_", value, flags=re.UNICODE).strip()
    return safe or "arquivo"


def normalize_name(value: str | None) -> str:
    """Normaliza nomes para comparações frouxas entre arquivos e datasources."""
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def decode_tableau_text(value: str | None) -> str | None:
    """Normaliza entidades XML frequentes nas fórmulas do Tableau."""
    if value is None:
        return None

    return (
        value.replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&#10;", "\n")
        .replace("&#13;", "\r")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
    )


def clean_brackets(value: str | None) -> str | None:
    """Remove colchetes do nome interno do Tableau, quando existirem."""
    if value is None:
        return None
    return value.replace("[", "").replace("]", "")


def clean_display_label(value: str | None) -> str | None:
    """Limpa ruídos visuais comuns em labels extraídos do XML."""
    if value is None:
        return None
    text = value.replace("\xa0", " ").strip()
    text = re.sub(r"\s{2,}", " ", text)
    text = re.sub(r"\s+([|,:;])", r"\1", text)
    text = re.sub(r"^\:+", "", text)
    return text.strip()


def is_color_like(value: str | None) -> bool:
    """Retorna True apenas para valores que parecem cores ou nomes de paleta úteis."""
    if not value:
        return False
    text = clean_display_label(value) or ""
    lower = text.lower()
    if re.fullmatch(r"#(?:[0-9a-fA-F]{3,8})", text):
        return True
    if lower.startswith(("rgb(", "rgba(", "hsl(", "hsla(")):
        return True
    if "palette" in lower:
        return True
    named_colors = {
        "black", "white", "red", "green", "blue", "yellow", "orange", "gray", "grey",
        "brown", "pink", "cyan", "magenta", "teal", "navy", "maroon", "olive", "lime",
        "silver", "gold", "beige", "transparent",
    }
    return lower in named_colors


def unique_ordered(values: list[Any]) -> list[Any]:
    """Remove duplicados preservando a ordem de aparição."""
    output = []
    seen = set()
    for value in values:
        marker = json.dumps(value, ensure_ascii=False, sort_keys=True) if isinstance(value, (dict, list)) else value
        if marker in seen:
            continue
        seen.add(marker)
        output.append(value)
    return output


def element_path_with_indices(root: ET.Element, target: ET.Element) -> str:
    """
    Monta um caminho XPath simples e legível até um elemento.
    O `xml.etree` não mantém ponteiro para o pai, então percorremos a árvore.
    """
    target_id = id(target)

    def walk(node: ET.Element, trail: list[str]) -> list[str] | None:
        if id(node) == target_id:
            return trail + [node.tag]

        child_counts: Counter[str] = Counter()
        for child in list(node):
            child_counts[child.tag] += 1
            siblings_same_tag = [item for item in list(node) if item.tag == child.tag]
            if len(siblings_same_tag) > 1:
                index = siblings_same_tag.index(child) + 1
                child_name = f"{child.tag}[{index}]"
            else:
                child_name = child.tag

            result = walk(child, trail + [node.tag, child_name])
            if result is not None:
                return result
        return None

    if id(root) == target_id:
        return f"/{root.tag}"

    result = walk(root, [])
    if result is None:
        return ""

    normalized = []
    skip_next_root = False
    for idx, part in enumerate(result):
        if idx == 0 and part == root.tag:
            normalized.append(part)
            skip_next_root = True
            continue
        if skip_next_root and part.startswith(f"{root.tag}/"):
            continue
        normalized.append(part)
    return "/" + "/".join(normalized)


class TableauDoc:
    """Carrega, extrai e documenta um workbook Tableau."""

    MAP_DEFINITIONS = [
        {
            "section": "workbook",
            "label": "Raiz do workbook",
            "xpath": ".",
            "json_path": "$.workbook",
            "description": "Metadados gerais do workbook, versão e build.",
        },
        {
            "section": "preferences",
            "label": "Preferências e paletas",
            "xpath": "./preferences",
            "json_path": "$.preferences",
            "description": "Preferências globais, incluindo paletas de cor customizadas.",
        },
        {
            "section": "styles",
            "label": "Estilos globais",
            "xpath": "./style/style-rule",
            "json_path": "$.styles.global_rules[*]",
            "description": "Regras de estilo globais, como fonte padrão e títulos.",
        },
        {
            "section": "datasources",
            "label": "Datasources",
            "xpath": "./datasources/datasource",
            "json_path": "$.datasources[*]",
            "description": "Fontes de dados, conexões, relações, colunas e metadados.",
        },
        {
            "section": "parameters",
            "label": "Parâmetros",
            "xpath": "./datasources/datasource[@name='Parameters']/column",
            "json_path": "$.parameters[*]",
            "description": "Parâmetros definidos no workbook.",
        },
        {
            "section": "calculations",
            "label": "Campos calculados",
            "xpath": ".//column[calculation] | .//calculation[@column]",
            "json_path": "$.calculations[*]",
            "description": "Campos calculados, fórmulas e dependências.",
        },
        {
            "section": "worksheets",
            "label": "Worksheets",
            "xpath": "./worksheets/worksheet",
            "json_path": "$.worksheets[*]",
            "description": "Planilhas, filtros, encodings, marks e estilos locais.",
        },
        {
            "section": "dashboards",
            "label": "Dashboards",
            "xpath": "./dashboards/dashboard",
            "json_path": "$.dashboards[*]",
            "description": "Painéis, zonas, objetos, filtros expostos e layout.",
        },
        {
            "section": "stories",
            "label": "Stories",
            "xpath": "./stories/story",
            "json_path": "$.stories[*]",
            "description": "Histórias do Tableau, quando presentes.",
        },
        {
            "section": "windows",
            "label": "Windows e cards",
            "xpath": "./windows/window",
            "json_path": "$.windows[*]",
            "description": "Estado da UI, cartões de filtro, cor, marks e legendas.",
        },
        {
            "section": "thumbnails",
            "label": "Miniaturas",
            "xpath": "./thumbnails/thumbnail",
            "json_path": "$.thumbnails[*]",
            "description": "Pré-visualizações embutidas em base64.",
        },
    ]

    def __init__(self, source_path: str | Path, output_format: str = "all") -> None:
        self.source_path = Path(source_path).expanduser().resolve()
        if not self.source_path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {self.source_path}")
        if self.source_path.suffix.lower() not in {".twb", ".twbx"}:
            raise ValueError("O arquivo informado deve ter extensão .twb ou .twbx.")

        self.output_format = output_format.lower()
        self.base_name = self.source_path.stem
        self.output_dir = DEFAULT_OUTPUT_ROOT / self.base_name
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.workbook_name_in_package: str | None = None
        self.workbook_xml_bytes: bytes | None = None
        self.package_manifest: list[dict[str, Any]] = []
        self.assets_written: list[str] = []
        self.caption_lookup: dict[str, str] = {}
        self.caption_lookup_by_clean: dict[str, str] = {}

        self.root = self._load_root_and_extract_contents()
        self.metadata = self._build_metadata()

    def _load_root_and_extract_contents(self) -> ET.Element:
        """Lê o XML principal e salva o conteúdo bruto do workbook no diretório de saída."""
        suffix = self.source_path.suffix.lower()

        if suffix == ".twb":
            self.workbook_name_in_package = self.source_path.name
            self.workbook_xml_bytes = self.source_path.read_bytes()
            extracted_dir = self.output_dir / "package_contents"
            if extracted_dir.exists():
                shutil.rmtree(extracted_dir)
            extracted_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self.source_path, extracted_dir / self.source_path.name)
            self.package_manifest.append(
                {
                    "path": self.source_path.name,
                    "kind": "workbook",
                    "size_bytes": self.source_path.stat().st_size,
                }
            )
            return ET.fromstring(self.workbook_xml_bytes)

        with zipfile.ZipFile(self.source_path, "r") as archive:
            extracted_dir = self.output_dir / "package_contents"
            if extracted_dir.exists():
                shutil.rmtree(extracted_dir)
            extracted_dir.mkdir(parents=True, exist_ok=True)
            for member in archive.infolist():
                destination = extracted_dir / member.filename
                if member.is_dir():
                    destination.mkdir(parents=True, exist_ok=True)
                    continue
                if member.filename.lower().endswith(".hyper"):
                    continue
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member, "r") as src, destination.open("wb") as dst:
                    shutil.copyfileobj(src, dst)

            for info in archive.infolist():
                self.package_manifest.append(
                    {
                        "path": info.filename,
                        "kind": self._classify_package_member(info.filename),
                        "size_bytes": info.file_size,
                    }
                )

            twb_names = sorted(
                name for name in archive.namelist() if name.lower().endswith(".twb")
            )
            if not twb_names:
                raise ValueError("O arquivo .twbx não contém um workbook .twb.")

            self.workbook_name_in_package = twb_names[0]
            self.workbook_xml_bytes = archive.read(twb_names[0])
            return ET.fromstring(self.workbook_xml_bytes)

    def _classify_package_member(self, package_path: str) -> str:
        """Classifica membros do pacote por tipo básico para documentação."""
        lower = package_path.lower()
        if lower.endswith(".twb"):
            return "workbook"
        if lower.endswith(".hyper"):
            return "extract"
        if lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg")):
            return "image"
        if lower.endswith((".csv", ".xlsx", ".xls", ".txt", ".json", ".xml")):
            return "data_or_support"
        return "other"

    def _build_metadata(self) -> dict[str, Any]:
        """Consolida o workbook em uma estrutura única para saída JSON/MD/Excel."""
        workbook_info = {
            "source_file_name": self.source_path.name,
            "source_file_path": str(self.source_path),
            "source_file_type": self.source_path.suffix.lower(),
            "source_file_last_modified": datetime.fromtimestamp(
                self.source_path.stat().st_mtime
            ).isoformat(timespec="seconds"),
            "output_directory": str(self.output_dir),
            "package_workbook_name": self.workbook_name_in_package,
            "attributes": dict(self.root.attrib),
            "repository_location": self._extract_repository_location(self.root),
        }

        preferences = self._extract_preferences()
        styles = self._extract_global_styles()
        datasources = self._extract_datasources()
        self._build_caption_lookup(datasources)
        parameters = self._extract_parameters(datasources)
        calculations = self._extract_calculations(datasources, parameters)
        worksheets = self._extract_worksheets()
        dashboards = self._extract_dashboards(worksheets)
        self._enrich_usage(parameters, calculations, worksheets, dashboards)
        stories = self._extract_stories()
        windows = self._extract_windows()
        thumbnails = self._extract_thumbnails()
        colors = self._extract_visual_tokens(token_type="color")
        fonts = self._extract_visual_tokens(token_type="font")

        return {
            "workbook": workbook_info,
            "package_manifest": self.package_manifest,
            "preferences": preferences,
            "styles": styles,
            "datasources": datasources,
            "parameters": parameters,
            "calculations": calculations,
            "worksheets": worksheets,
            "dashboards": dashboards,
            "stories": stories,
            "windows": windows,
            "thumbnails": thumbnails,
            "visual_tokens": {
                "colors": colors,
                "fonts": fonts,
            },
            "summary": {
                "datasource_count": len(datasources),
                "parameter_count": len(parameters),
                "calculation_count": len(calculations),
                "worksheet_count": len(worksheets),
                "dashboard_count": len(dashboards),
                "story_count": len(stories),
                "window_count": len(windows),
                "thumbnail_count": len(thumbnails),
                "package_member_count": len(self.package_manifest),
            },
        }

    def _build_caption_lookup(self, datasources: list[dict[str, Any]]) -> None:
        """Monta um mapa global entre nomes internos e captions amigáveis."""
        self.caption_lookup = {}
        self.caption_lookup_by_clean = {}
        for datasource in datasources:
            for column in datasource.get("columns", []):
                name = column.get("name")
                caption = column.get("caption") or clean_brackets(name)
                if name and caption:
                    self.caption_lookup[name] = caption
                    self.caption_lookup_by_clean[clean_brackets(name) or ""] = caption
            for record in datasource.get("metadata_records", []):
                local_name = record.get("local_name")
                caption = record.get("caption") or clean_brackets(local_name)
                if local_name and caption:
                    self.caption_lookup[local_name] = caption
                    self.caption_lookup_by_clean[clean_brackets(local_name) or ""] = caption

    def _extract_repository_location(self, element: ET.Element) -> dict[str, Any] | None:
        repository = element.find("./repository-location")
        return dict(repository.attrib) if repository is not None else None

    def _extract_preferences(self) -> dict[str, Any]:
        preferences = self.root.find("./preferences")
        if preferences is None:
            return {"preferences": [], "color_palettes": []}

        pref_rows = [dict(pref.attrib) for pref in preferences.findall("./preference")]
        palettes = []
        for palette in preferences.findall("./color-palette"):
            palettes.append(
                {
                    "name": palette.get("name"),
                    "type": palette.get("type"),
                    "custom": palette.get("custom"),
                    "colors": [color.text for color in palette.findall("./color") if color.text],
                }
            )
        return {"preferences": pref_rows, "color_palettes": palettes}

    def _extract_global_styles(self) -> dict[str, Any]:
        style = self.root.find("./style")
        return {"global_rules": self._parse_style_rules(style) if style is not None else []}

    def _extract_datasources(self) -> list[dict[str, Any]]:
        datasources = []
        for datasource in self.root.findall("./datasources/datasource"):
            columns = []
            for column in datasource.findall("./column"):
                columns.append(self._parse_column(column))

            metadata_records = []
            for record in datasource.findall(".//metadata-records/metadata-record"):
                metadata_records.append(self._parse_metadata_record(record))

            calculations = []
            for calc in datasource.findall(".//connection/calculations/calculation"):
                calculations.append(
                    {
                        "column": calc.get("column"),
                        "formula": decode_tableau_text(calc.get("formula")),
                    }
                )

            connections = []
            for connection in datasource.findall("./connection"):
                connections.append(self._parse_connection(connection))

            datasource_info = {
                "name": datasource.get("name"),
                "caption": datasource.get("caption"),
                "version": datasource.get("version"),
                "inline": datasource.get("inline"),
                "hasconnection": datasource.get("hasconnection"),
                "repository_location": self._extract_repository_location(datasource),
                "connections": connections,
                "columns": columns,
                "metadata_records": metadata_records,
                "connection_calculations": calculations,
                "aliases_enabled": datasource.find("./aliases") is not None and datasource.find("./aliases").get("enabled"),
                "object_count": len(datasource.findall("./object-graph/object")),
            }
            datasources.append(datasource_info)
        return datasources

    def _parse_connection(self, connection: ET.Element) -> dict[str, Any]:
        relation_rows = []
        for relation in connection.findall(".//relation"):
            relation_rows.append(dict(relation.attrib))

        named_connections = []
        for named in connection.findall("./named-connections/named-connection"):
            named_connections.append(dict(named.attrib))

        return {
            "attributes": dict(connection.attrib),
            "relations": relation_rows,
            "named_connections": named_connections,
        }

    def _parse_column(self, column: ET.Element) -> dict[str, Any]:
        aliases = [
            {"key": alias.get("key"), "value": alias.get("value")}
            for alias in column.findall("./aliases/alias")
        ]
        members = [member.get("value") for member in column.findall("./members/member") if member.get("value")]
        bins = []
        for bin_element in column.findall("./calculation/bin"):
            bins.append(
                {
                    "value": bin_element.get("value"),
                    "default_name": bin_element.get("default-name"),
                    "members": [value.text for value in bin_element.findall("./value") if value.text],
                }
            )

        calculation = column.find("./calculation")
        return {
            "name": column.get("name"),
            "caption": column.get("caption"),
            "role": column.get("role"),
            "datatype": column.get("datatype"),
            "type": column.get("type"),
            "default_type": column.get("default-type"),
            "default_format": column.get("default-format"),
            "aggregation": column.get("aggregation"),
            "hidden": column.get("hidden"),
            "value": column.get("value"),
            "param_domain_type": column.get("param-domain-type"),
            "semantic_role": column.get("semantic-role"),
            "source_field": column.get("source-field"),
            "aliases": aliases,
            "members": members,
            "bins": bins,
            "formula": decode_tableau_text(calculation.get("formula")) if calculation is not None else None,
            "calculation_class": calculation.get("class") if calculation is not None else None,
        }

    def _parse_metadata_record(self, record: ET.Element) -> dict[str, Any]:
        attributes = {}
        for attribute in record.findall("./attributes/attribute"):
            key = attribute.get("name")
            if key:
                attributes[key] = decode_tableau_text(attribute.text or "")

        return {
            "class": record.get("class"),
            "caption": record.findtext("./caption"),
            "local_name": record.findtext("./local-name"),
            "remote_name": record.findtext("./remote-name"),
            "parent_name": record.findtext("./parent-name"),
            "local_type": record.findtext("./local-type"),
            "aggregation": record.findtext("./aggregation"),
            "attributes": attributes,
        }

    def _extract_parameters(self, datasources: list[dict[str, Any]]) -> list[dict[str, Any]]:
        for datasource in datasources:
            if datasource.get("name") == "Parameters":
                return datasource["columns"]
        return []

    def _extract_calculations(
        self,
        datasources: list[dict[str, Any]],
        parameters: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        friendly_names = {}
        for parameter in parameters:
            if parameter.get("name") and parameter.get("caption"):
                friendly_names[parameter["name"]] = parameter["caption"]

        records = []
        seen = set()

        for datasource in datasources:
            datasource_name = datasource.get("caption") or datasource.get("name")
            if datasource.get("name") == "Parameters":
                continue
            for column in datasource.get("columns", []):
                if not column.get("formula"):
                    continue
                if column.get("name") and column.get("caption"):
                    friendly_names[column["name"]] = column["caption"]

                key = ("column", datasource_name, column.get("name"), column.get("formula"))
                if key in seen:
                    continue
                seen.add(key)
                records.append(
                    {
                        "datasource": datasource_name,
                        "origin": "column",
                        "name": column.get("name"),
                        "caption": column.get("caption") or clean_brackets(column.get("name")),
                        "role": column.get("role"),
                        "datatype": column.get("datatype"),
                        "type": column.get("type"),
                        "formula": column.get("formula"),
                    }
                )

            for calc in datasource.get("connection_calculations", []):
                key = ("connection", datasource_name, calc.get("column"), calc.get("formula"))
                if key in seen:
                    continue
                seen.add(key)
                records.append(
                    {
                        "datasource": datasource_name,
                        "origin": "connection.calculations",
                        "name": calc.get("column"),
                        "caption": clean_brackets(calc.get("column")),
                        "role": None,
                        "datatype": None,
                        "type": None,
                        "formula": calc.get("formula"),
                    }
                )

            for record in datasource.get("metadata_records", []):
                formula = record.get("attributes", {}).get("formula")
                if not formula:
                    continue
                key = ("metadata", datasource_name, record.get("local_name"), formula)
                if key in seen:
                    continue
                seen.add(key)
                records.append(
                    {
                        "datasource": datasource_name,
                        "origin": "metadata-record",
                        "name": record.get("local_name"),
                        "caption": record.get("caption") or clean_brackets(record.get("local_name")),
                        "role": None,
                        "datatype": record.get("local_type"),
                        "type": record.get("class"),
                        "formula": formula.strip('"'),
                    }
                )

        for record in records:
            if record.get("name") and record.get("caption"):
                friendly_names[record["name"]] = record["caption"]
                cleaned_name = clean_brackets(record["name"])
                if cleaned_name:
                    friendly_names[f"[{cleaned_name}]"] = record["caption"]
            record["formula_pretty"] = self._replace_internal_names(record.get("formula"), friendly_names)
            record["codigo"] = record["formula_pretty"] or record.get("formula")

        by_datasource = {}
        for record in records:
            by_datasource.setdefault(record["datasource"], []).append(record)

        for datasource_records in by_datasource.values():
            caption_by_name = {}
            for record in datasource_records:
                if record.get("name") and record.get("caption"):
                    caption_by_name[record["name"]] = record["caption"]

            for record in datasource_records:
                impacts = []
                target = f"[{record['caption']}]"
                for candidate in datasource_records:
                    if candidate is record:
                        continue
                    formula = candidate.get("formula_pretty") or ""
                    if target in formula:
                        impacts.append(candidate["caption"])
                record["impacts"] = sorted(unique_ordered([impact for impact in impacts if impact]), key=str.lower)

        return records

    def _replace_internal_names(self, formula: str | None, friendly_names: dict[str, str]) -> str | None:
        if not formula or not friendly_names:
            return formula

        pattern = re.compile(
            "|".join(re.escape(key) for key in sorted(friendly_names.keys(), key=len, reverse=True))
        )
        return pattern.sub(lambda match: f"[{friendly_names[match.group(0)]}]", formula)

    def _enrich_usage(
        self,
        parameters: list[dict[str, Any]],
        calculations: list[dict[str, Any]],
        worksheets: list[dict[str, Any]],
        dashboards: list[dict[str, Any]],
    ) -> None:
        """Mapeia uso de parâmetros e cálculos em worksheets e dashboards."""
        worksheet_xml = {
            worksheet.get("name"): ET.tostring(element, encoding="unicode")
            for worksheet, element in zip(
                worksheets,
                self.root.findall("./worksheets/worksheet"),
            )
        }
        dashboard_members = {
            dashboard["name"]: set(dashboard.get("worksheet_members", []))
            for dashboard in dashboards
        }

        for parameter in parameters:
            used_in_worksheets = sorted(
                name
                for name, xml in worksheet_xml.items()
                if self._xml_contains_reference(xml, parameter.get("name"), parameter.get("caption"))
            )
            used_in_dashboards = sorted(
                dashboard_name
                for dashboard_name, members in dashboard_members.items()
                if members.intersection(used_in_worksheets)
            )
            parameter["used_in_worksheets"] = used_in_worksheets
            parameter["used_in_dashboards"] = used_in_dashboards
            parameter["is_used"] = bool(used_in_worksheets or used_in_dashboards)
            parameter["source_field_details"] = self._parse_source_field_reference(parameter.get("source_field"))

        for calculation in calculations:
            used_in_worksheets = sorted(
                name
                for name, xml in worksheet_xml.items()
                if self._xml_contains_reference(xml, calculation.get("name"), calculation.get("caption"))
            )
            used_in_dashboards = sorted(
                dashboard_name
                for dashboard_name, members in dashboard_members.items()
                if members.intersection(used_in_worksheets)
            )
            calculation["used_in_worksheets"] = used_in_worksheets
            calculation["used_in_dashboards"] = used_in_dashboards
            calculation["is_used"] = bool(used_in_worksheets or used_in_dashboards)

    def _xml_contains_reference(self, xml_text: str, internal_name: str | None, caption: str | None) -> bool:
        """Busca referências de um campo no XML de uma worksheet."""
        candidates = []
        if internal_name:
            candidates.append(internal_name)
            clean_name = clean_brackets(internal_name)
            if clean_name:
                candidates.append(f"[{clean_name}]")
        if caption:
            candidates.append(f"[{caption}]")
        return any(candidate and candidate in xml_text for candidate in unique_ordered(candidates))

    def _parse_source_field_reference(self, value: str | None) -> dict[str, str] | None:
        """Quebra a referência source-field em datasource e campo, quando possível."""
        if not value:
            return None
        match = re.match(r"^\[(?P<datasource>.+?)\]\.\[(?P<field>.+?)\]$", value)
        if not match:
            return {"raw": value}
        return {
            "datasource": match.group("datasource"),
            "field": match.group("field"),
            "raw": value,
        }

    def _extract_hyper_structures(self, datasources: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """
        Documenta a estrutura dos extracts a partir do metadata do workbook.

        Observação: isso é inferido do XML do Tableau. O ambiente atual não
        possui a Hyper API instalada para introspecção binária direta.
        """
        hyper_files = [item["path"] for item in self.package_manifest if item["kind"] == "extract"]
        structures = []

        for datasource in datasources:
            if datasource.get("name") == "Parameters":
                continue
            has_extract = any(
                relation.get("table", "").startswith("[Extract].")
                for connection in datasource.get("connections", [])
                for relation in connection.get("relations", [])
            )
            if not has_extract:
                continue

            matched_hyper = self._match_hyper_file(datasource, hyper_files)
            fields = []
            for column in datasource.get("columns", []):
                field_name = column.get("caption") or clean_brackets(column.get("name"))
                if not field_name:
                    continue
                fields.append(
                    {
                        "field": field_name,
                        "internal_name": column.get("name"),
                        "datatype": column.get("datatype"),
                        "format": column.get("default_format") or column.get("default_type"),
                        "role": self._friendly_role(column),
                        "source": "workbook_metadata",
                    }
                )

            fields = sorted(fields, key=lambda item: (item["field"] or "").lower())
            structures.append(
                {
                    "hyper_file": matched_hyper,
                    "datasource": datasource.get("caption") or datasource.get("name"),
                    "note": "Estrutura inferida do metadata do workbook; o binário .hyper não foi copiado para data.",
                    "field_count": len(fields),
                    "fields": fields,
                }
            )

        return structures

    def _match_hyper_file(self, datasource: dict[str, Any], hyper_files: list[str]) -> str | None:
        """Relaciona um datasource a um arquivo .hyper do pacote por similaridade de nome."""
        datasource_name = normalize_name(datasource.get("caption") or datasource.get("name"))
        best_match = None
        best_score = -1
        for hyper_file in hyper_files:
            hyper_name = normalize_name(Path(hyper_file).stem)
            score = len(set(re.findall(r"[a-z0-9]+", datasource_name)).intersection(re.findall(r"[a-z0-9]+", hyper_name)))
            if datasource_name and datasource_name in hyper_name:
                score += 10
            if score > best_score:
                best_match = hyper_file
                best_score = score
        return best_match

    def _friendly_role(self, column: dict[str, Any]) -> str:
        """Traduz role/tipo do Tableau para uma forma simples no relatório."""
        role = (column.get("role") or "").lower()
        if role == "measure":
            return "measure"
        if role == "dimension":
            return "dimension"
        default_type = (column.get("default_type") or "").lower()
        if default_type in {"quantitative", "ordinal"}:
            return "measure"
        return "dimension"

    def _extract_worksheets(self) -> list[dict[str, Any]]:
        worksheets = []
        for worksheet in self.root.findall("./worksheets/worksheet"):
            worksheet_name = worksheet.get("name")
            formatted_title = self._extract_formatted_text(worksheet.find("./layout-options/title/formatted-text"))
            datasource_dependencies = [
                dependency.get("datasource")
                for dependency in worksheet.findall(".//datasource-dependencies")
                if dependency.get("datasource")
            ]
            filters = [self._parse_filter(filter_element) for filter_element in worksheet.findall(".//filter")]
            encodings = self._parse_encodings(worksheet)
            marks = [mark.get("class") for mark in worksheet.findall(".//mark") if mark.get("class")]
            style_rules = self._parse_style_rules(worksheet)

            worksheets.append(
                {
                    "name": worksheet_name,
                    "title": formatted_title,
                    "datasource_dependencies": unique_ordered(datasource_dependencies),
                    "filters": filters,
                    "encodings": encodings,
                    "marks": unique_ordered(marks),
                    "style_rules": style_rules,
                    "colors_used": unique_ordered(self._collect_colors(worksheet)),
                    "fonts_used": unique_ordered(self._collect_fonts(worksheet)),
                    "referenced_columns": unique_ordered(self._collect_columns_from_dependencies(worksheet)),
                    "shelf_columns": [column.text for column in worksheet.findall(".//slices/column") if column.text],
                }
            )
        return worksheets

    def _extract_dashboards(self, worksheets: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        dashboards = []
        worksheets = worksheets or []
        worksheet_names = {worksheet.get("name") for worksheet in self.root.findall("./worksheets/worksheet")}

        for dashboard in self.root.findall("./dashboards/dashboard"):
            zones = self._parse_zones(dashboard.find("./zones"))
            dashboard_name = dashboard.get("name")
            sheets = []
            exposed_filters = []
            for zone in zones:
                if zone.get("name") in worksheet_names:
                    sheets.append(zone["name"])
                if zone.get("type_v2") == "filter":
                    exposed_filters.append(
                        {
                            "zone_id": zone.get("id"),
                            "worksheet_or_dashboard": zone.get("name"),
                            "param": zone.get("param"),
                            "mode": zone.get("mode"),
                            "values": zone.get("values"),
                            "field_label": self._humanize_field_reference(zone.get("param")),
                            "is_context": False,
                        }
                    )

            device_layouts = []
            for layout in dashboard.findall("./devicelayouts/devicelayout"):
                device_layouts.append(
                    {
                        "name": layout.get("name"),
                        "auto_generated": layout.get("auto-generated"),
                        "zones": self._parse_zones(layout.find("./zones")),
                    }
                )

            worksheet_filters = []
            for worksheet_name in unique_ordered([sheet for sheet in sheets if sheet]):
                worksheet = next((item for item in worksheets if item.get("name") == worksheet_name), None)
                if worksheet:
                    worksheet_filters.extend(worksheet.get("filters", []))

            combined_filters = []
            for item in exposed_filters + worksheet_filters:
                field_label = item.get("field_label") or self._humanize_field_reference(item.get("column") or item.get("param"))
                combined_filters.append(
                    {
                        "field_label": field_label,
                        "class": item.get("class"),
                        "is_context": bool(item.get("is_context")),
                        "visibility": "exposto no painel" if "zone_id" in item else "interno da planilha",
                    }
                )
            combined_filters = sorted(
                unique_ordered(combined_filters),
                key=lambda item: (item.get("field_label") or "").lower(),
            )

            dashboards.append(
                {
                    "name": dashboard_name,
                    "attributes": dict(dashboard.attrib),
                    "datasource_dependencies": unique_ordered(
                        [
                            dependency.get("datasource")
                            for dependency in dashboard.findall("./datasource-dependencies")
                            if dependency.get("datasource")
                        ]
                    ),
                    "zones": zones,
                    "worksheet_members": unique_ordered([sheet for sheet in sheets if sheet]),
                    "filters_exposed": exposed_filters,
                    "filters_used": combined_filters,
                    "device_layouts": device_layouts,
                    "colors_used": unique_ordered(self._collect_colors(dashboard)),
                    "fonts_used": unique_ordered(self._collect_fonts(dashboard)),
                }
            )
        return dashboards

    def _extract_stories(self) -> list[dict[str, Any]]:
        stories = []
        for story in self.root.findall("./stories/story"):
            stories.append(
                {
                    "name": story.get("name"),
                    "attributes": dict(story.attrib),
                    "style_rules": self._parse_style_rules(story),
                    "colors_used": unique_ordered(self._collect_colors(story)),
                    "fonts_used": unique_ordered(self._collect_fonts(story)),
                }
            )
        return stories

    def _extract_windows(self) -> list[dict[str, Any]]:
        windows = []
        for window in self.root.findall("./windows/window"):
            cards = []
            for card in window.findall(".//card"):
                cards.append(dict(card.attrib))
            windows.append(
                {
                    "class": window.get("class"),
                    "name": window.get("name"),
                    "maximized": window.get("maximized"),
                    "cards": cards,
                }
            )
        return windows

    def _extract_thumbnails(self) -> list[dict[str, Any]]:
        thumbnails = []
        thumbnails_dir = self.output_dir / "thumbnails"
        if thumbnails_dir.exists():
            shutil.rmtree(thumbnails_dir)
        thumbnails_dir.mkdir(parents=True, exist_ok=True)

        for index, thumbnail in enumerate(self.root.findall("./thumbnails/thumbnail"), start=1):
            name = thumbnail.get("name") or f"thumbnail_{index}"
            file_name = f"{sanitize_filename(name)}.png"
            output_path = thumbnails_dir / file_name
            if thumbnail.text:
                output_path.write_bytes(base64.decodebytes(thumbnail.text.encode()))
                self.assets_written.append(str(output_path))
            thumbnails.append(
                {
                    "name": name,
                    "output_path": str(output_path),
                    "has_embedded_data": bool(thumbnail.text),
                }
            )
        return thumbnails

    def _extract_visual_tokens(self, token_type: str) -> list[dict[str, Any]]:
        """Varre o XML inteiro em busca de cores ou fontes com contexto e XPath."""
        results = []
        for element in self.root.iter():
            if token_type == "color":
                format_attr = element.get("attr")
                format_value = element.get("value")
                run_color = element.get("fontcolor")
                if format_attr and "color" in format_attr and format_value:
                    if is_color_like(format_value):
                        results.append(
                            {
                                "value": clean_display_label(format_value),
                                "context_tag": element.tag,
                                "attribute": format_attr,
                                "xpath": self._best_effort_xpath(element),
                            }
                        )
                if run_color:
                    if is_color_like(run_color):
                        results.append(
                            {
                                "value": clean_display_label(run_color),
                                "context_tag": element.tag,
                                "attribute": "fontcolor",
                                "xpath": self._best_effort_xpath(element),
                            }
                        )
            else:
                if element.get("attr") == "font-family" and element.get("value"):
                    results.append(
                        {
                            "value": element.get("value"),
                            "context_tag": element.tag,
                            "attribute": "font-family",
                            "xpath": self._best_effort_xpath(element),
                        }
                    )
                for attr_name in ("fontname", "font-family"):
                    if element.get(attr_name):
                        results.append(
                            {
                                "value": element.get(attr_name),
                                "context_tag": element.tag,
                                "attribute": attr_name,
                                "xpath": self._best_effort_xpath(element),
                            }
                        )
        return unique_ordered(results)

    def _parse_filter(self, filter_element: ET.Element) -> dict[str, Any]:
        members = []
        for groupfilter in filter_element.findall(".//groupfilter"):
            entry = {key: value for key, value in groupfilter.attrib.items()}
            if groupfilter.text and groupfilter.text.strip():
                entry["text"] = groupfilter.text.strip()
            members.append(entry)

        column_ref = filter_element.get("column")
        return {
            "class": filter_element.get("class"),
            "column": column_ref,
            "field_label": self._humanize_field_reference(column_ref),
            "filter_group": filter_element.get("filter-group"),
            "included_values": filter_element.get("included-values"),
            "is_context": self._is_context_filter(filter_element),
            "groupfilters": members,
        }

    def _is_context_filter(self, filter_element: ET.Element) -> bool:
        """Tenta identificar marcadores de contexto associados a um filtro."""
        serialized = ET.tostring(filter_element, encoding="unicode")
        return "context" in serialized.lower()

    def _humanize_field_reference(self, value: str | None) -> str | None:
        """Converte referências internas do Tableau em um nome mais legível."""
        if not value:
            return None
        original = value

        direct_candidates = re.findall(r"\[[^\]]+\]", original)
        for candidate in direct_candidates:
            if candidate in self.caption_lookup:
                return clean_display_label(self.caption_lookup[candidate])
            candidate_clean = clean_brackets(candidate)
            if candidate_clean in self.caption_lookup_by_clean:
                return clean_display_label(self.caption_lookup_by_clean[candidate_clean])

        text = original
        text = re.sub(r"^\[[^\]]+\]\.", "", text)
        text = text.replace("[", "").replace("]", "")
        text = re.sub(r"^(sum|avg|min|max|cnt|usr|none):", "", text)
        text = re.sub(r":[a-z]{1,3}$", "", text)
        text = re.sub(r"^\:+", "", text)

        if text in self.caption_lookup_by_clean:
            return clean_display_label(self.caption_lookup_by_clean[text])

        calc_match = re.search(r"(Calculation_\d+)", text)
        if calc_match and calc_match.group(1) in self.caption_lookup_by_clean:
            return clean_display_label(self.caption_lookup_by_clean[calc_match.group(1)])

        special_cases = {
            "Measure Names": "Measure Names",
            "Measure Values": "Measure Values",
            ":Measure Names": "Measure Names",
            ":Measure Values": "Measure Values",
        }
        if text in special_cases:
            return special_cases[text]
        if "Measure Names" in text:
            return "Measure Names"
        if "Measure Values" in text:
            return "Measure Values"

        text = re.sub(r"\s+\(\s*copy\s*\)\s*$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\(\s*cópia\s*\)\s*$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\(\s*local copy\s*\)\s*$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\(\s*local\s*\)\s*$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\|\s+snapshot$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\|\s+current year$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\|\s+next year$", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s+\(\s*prod_base_[^)]+\)$", "", text, flags=re.IGNORECASE)

        return clean_display_label(text)

    def _parse_encodings(self, element: ET.Element) -> list[dict[str, Any]]:
        encodings = []
        for encoding in element.findall(".//encodings/*"):
            record = {"tag": encoding.tag}
            record.update(encoding.attrib)
            encodings.append(record)
        for encoding in element.findall(".//style-rule/encoding"):
            record = {"tag": "style-encoding"}
            record.update(encoding.attrib)
            encodings.append(record)
        return encodings

    def _parse_style_rules(self, parent: ET.Element | None) -> list[dict[str, Any]]:
        if parent is None:
            return []
        rules = []
        for rule in parent.findall(".//style-rule"):
            formats = []
            for fmt in rule.findall("./format"):
                formats.append(dict(fmt.attrib))
            encodings = []
            for encoding in rule.findall("./encoding"):
                encodings.append(dict(encoding.attrib))
            rules.append(
                {
                    "element": rule.get("element"),
                    "formats": formats,
                    "encodings": encodings,
                }
            )
        return rules

    def _extract_formatted_text(self, formatted_text: ET.Element | None) -> list[dict[str, Any]]:
        if formatted_text is None:
            return []
        runs = []
        for run in formatted_text.findall("./run"):
            runs.append(
                {
                    "text": "".join(run.itertext()).strip(),
                    "attributes": dict(run.attrib),
                }
            )
        return runs

    def _collect_colors(self, element: ET.Element) -> list[str]:
        colors = []
        for item in element.iter():
            if item.get("attr") and "color" in item.get("attr", "") and item.get("value"):
                if is_color_like(item.get("value")):
                    colors.append(clean_display_label(item.get("value")))
            if item.get("fontcolor"):
                if is_color_like(item.get("fontcolor")):
                    colors.append(clean_display_label(item.get("fontcolor")))
        return colors

    def _collect_fonts(self, element: ET.Element) -> list[str]:
        fonts = []
        for item in element.iter():
            if item.get("attr") == "font-family" and item.get("value"):
                fonts.append(item.get("value"))
            if item.get("fontname"):
                fonts.append(item.get("fontname"))
        return fonts

    def _collect_columns_from_dependencies(self, element: ET.Element) -> list[str]:
        columns = []
        for column in element.findall(".//datasource-dependencies/column"):
            columns.append(column.get("caption") or clean_brackets(column.get("name")))
        return [item for item in columns if item]

    def _parse_zones(self, zones_root: ET.Element | None) -> list[dict[str, Any]]:
        if zones_root is None:
            return []

        parsed = []

        def visit(zone: ET.Element, parent_id: str | None = None) -> None:
            zone_style = []
            for fmt in zone.findall("./zone-style/format"):
                zone_style.append(dict(fmt.attrib))

            record = {
                "id": zone.get("id"),
                "name": zone.get("name"),
                "type_v2": zone.get("type-v2"),
                "param": zone.get("param"),
                "mode": zone.get("mode"),
                "values": zone.get("values"),
                "x": zone.get("x"),
                "y": zone.get("y"),
                "w": zone.get("w"),
                "h": zone.get("h"),
                "parent_id": parent_id,
                "style": zone_style,
            }
            parsed.append(record)

            for child_zone in zone.findall("./zone"):
                visit(child_zone, zone.get("id"))

        for zone in zones_root.findall("./zone"):
            visit(zone)
        return parsed

    def _best_effort_xpath(self, element: ET.Element) -> str:
        try:
            return element_path_with_indices(self.root, element)
        except Exception:
            return ""

    def generate_xpath_json_map(self) -> tuple[Path, Path]:
        """Gera o documento humano-legível e a versão JSON do mapa XPath/JSON."""
        map_rows = []
        for definition in self.MAP_DEFINITIONS:
            count = self._count_xpath_matches(definition["xpath"])
            map_rows.append(
                {
                    **definition,
                    "match_count": count,
                }
            )

        map_json_path = self.output_dir / "mapa_XPath_JSON.json"
        map_md_path = self.output_dir / "mapa_XPath_JSON.md"
        map_json_path.write_text(
            json.dumps(
                {
                    "source_file": str(self.source_path),
                    "output_directory": str(self.output_dir),
                    "map_entries": map_rows,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        lines = [
            f"# mapa_XPath_JSON - {self.source_path.name}",
            "",
            "## Como Ler",
            "",
            "- `XPath`: local esperado no XML do Tableau.",
            "- `JSON Path`: local correspondente na saída JSON gerada pelo script.",
            "- `Ocorrências`: quantidade encontrada no arquivo analisado.",
            "",
            "## Mapa",
            "",
            "| Seção | Label | XPath | JSON Path | Ocorrências | Descrição |",
            "|---|---|---|---|---:|---|",
        ]

        for row in map_rows:
            lines.append(
                f"| {row['section']} | {row['label']} | `{row['xpath']}` | `{row['json_path']}` | {row['match_count']} | {row['description']} |"
            )

        map_md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return map_md_path, map_json_path

    def _count_xpath_matches(self, xpath: str) -> int:
        if xpath == ".":
            return 1
        if "|" in xpath:
            parts = [part.strip() for part in xpath.split("|")]
            count = 0
            for part in parts:
                count += len(self.root.findall(part))
            return count
        return len(self.root.findall(xpath))

    def write_outputs(self) -> list[Path]:
        """Escreve o mapa e os artefatos solicitados para o diretório de saída."""
        written_files: list[Path] = []

        workbook_xml_path = self.output_dir / f"{self.base_name}.xml"
        workbook_xml_path.write_bytes(self.workbook_xml_bytes or b"")
        written_files.append(workbook_xml_path)

        map_md_path, map_json_path = self.generate_xpath_json_map()
        written_files.extend([map_md_path, map_json_path])

        if self.output_format in {"all", "json"}:
            written_files.append(self._write_json())
        if self.output_format in {"all", "markdown"}:
            written_files.append(self._write_markdown())
        if self.output_format in {"all", "excel"}:
            written_files.append(self._write_excel())

        manifest_path = self.output_dir / f"{self.base_name}_manifest.json"
        manifest_path.write_text(
            json.dumps(
                {
                    "source_file": str(self.source_path),
                    "output_directory": str(self.output_dir),
                    "generated_files": [str(path) for path in written_files],
                    "package_manifest": self.package_manifest,
                    "thumbnails_written": self.metadata["thumbnails"],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        written_files.append(manifest_path)
        self._cleanup_temporary_outputs()
        return written_files

    def _cleanup_temporary_outputs(self) -> None:
        """
        Remove artefatos temporários usados apenas durante o processamento.

        Mantém os arquivos finais de documentação e miniaturas, removendo
        diretórios intermediários temporários.
        """
        for name in TEMPORARY_OUTPUT_NAMES:
            path = self.output_dir / name
            if not path.exists():
                continue
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            else:
                try:
                    path.unlink()
                except FileNotFoundError:
                    pass

    def _write_json(self) -> Path:
        output_path = self.output_dir / f"{self.base_name}.json"
        output_path.write_text(
            json.dumps(self.metadata, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return output_path

    def _write_markdown(self) -> Path:
        output_path = self.output_dir / f"{self.base_name}.md"
        lines = [
            f"# Documentação do Workbook Tableau - {self.source_path.name}",
            "",
            "## Resumo",
            "",
            f"- Caminho de origem: `{self.source_path}`",
            f"- Tipo do arquivo: `{self.source_path.suffix.lower()}`",
            f"- Última alteração do arquivo: `{self.metadata['workbook']['source_file_last_modified']}`",
            f"- Diretório de saída: `{self.output_dir}`",
            f"- Datasources: {self.metadata['summary']['datasource_count']}",
            f"- Parâmetros: {self.metadata['summary']['parameter_count']}",
            f"- Campos calculados: {self.metadata['summary']['calculation_count']}",
            f"- Worksheets: {self.metadata['summary']['worksheet_count']}",
            f"- Dashboards: {self.metadata['summary']['dashboard_count']}",
            "",
            "## Workbook",
            "",
        ]

        for key, value in self.metadata["workbook"]["attributes"].items():
            lines.append(f"- {key}: `{value}`")
        if self.metadata["workbook"]["repository_location"]:
            lines.append("")
            lines.append("### Repository Location")
            lines.append("")
            for key, value in self.metadata["workbook"]["repository_location"].items():
                lines.append(f"- {key}: `{value}`")

        lines.extend(self._build_datasources_markdown())
        lines.extend(self._build_dashboards_markdown())
        lines.extend(self._build_visual_tokens_markdown())
        lines.extend(self._build_preferences_markdown())
        lines.extend(self._build_parameters_markdown())
        lines.extend(self._build_calculations_markdown())

        output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return output_path

    def _build_datasources_markdown(self) -> list[str]:
        lines = ["", "## Fontes de Dados", ""]
        datasources = self.metadata["datasources"]
        if not datasources:
            lines.append("- Nenhuma fonte de dados identificada.")
            return lines

        for datasource in datasources:
            name = datasource.get("caption") or datasource.get("name") or "(sem nome)"
            repo = datasource.get("repository_location") or {}
            is_published = bool(repo)
            lines.append(f"### {name}")
            lines.append("")
            lines.append(f"- Nome interno: `{datasource.get('name')}`")
            lines.append(f"- Versão: `{datasource.get('version')}`")
            lines.append(f"- Publicada: `{'sim' if is_published else 'não'}`")
            if is_published:
                lines.append(f"- Site: `{repo.get('site', '-')}`")
                lines.append(f"- Path: `{repo.get('path', '-')}`")
                lines.append(f"- ID: `{repo.get('id', '-')}`")
                lines.append(f"- Revision: `{repo.get('revision', '-')}`")
            lines.append(f"- Quantidade de campos mapeados: {len(datasource.get('columns', []))}")
            lines.append("")
        return lines

    def _build_preferences_markdown(self) -> list[str]:
        lines = ["", "## Preferências e Paletas", ""]
        palettes = self.metadata["preferences"]["color_palettes"]
        prefs = self.metadata["preferences"]["preferences"]

        if not prefs and not palettes:
            lines.append("- Nenhuma preferência global encontrada.")
            return lines

        for pref in prefs:
            lines.append(f"- Preferência `{pref.get('name')}` = `{pref.get('value')}`")

        if palettes:
            lines.append("")
            lines.append("### Paletas de Cor")
            lines.append("")
            for palette in palettes:
                colors = ", ".join(palette.get("colors", [])) or "-"
                lines.append(f"- `{palette.get('name')}` ({palette.get('type')}) -> {colors}")
        return lines

    def _build_parameters_markdown(self) -> list[str]:
        lines = ["", "## Parâmetros", ""]
        parameters = self.metadata["parameters"]
        if not parameters:
            lines.append("- Nenhum parâmetro encontrado.")
            return lines

        for parameter in parameters:
            name = parameter.get("caption") or clean_brackets(parameter.get("name")) or "(sem nome)"
            members = ", ".join(parameter.get("members", [])) or "-"
            lines.append(f"### {name}")
            lines.append("")
            lines.append(f"- Nome interno: `{parameter.get('name')}`")
            lines.append(f"- Datatype: `{parameter.get('datatype')}`")
            lines.append(f"- Type: `{parameter.get('type')}`")
            lines.append(f"- Role: `{parameter.get('role')}`")
            lines.append(f"- Domínio: `{parameter.get('param_domain_type')}`")
            lines.append(f"- Valor atual/default: `{parameter.get('value')}`")
            lines.append(f"- Membros: {members}")
            source_field = parameter.get("source_field_details")
            if source_field:
                if source_field.get("datasource") and source_field.get("field"):
                    lines.append(f"- Preenchido por campo da fonte: `{source_field['datasource']}` -> `{source_field['field']}`")
                else:
                    lines.append(f"- Preenchido por campo da fonte: `{source_field.get('raw')}`")
            else:
                lines.append("- Preenchido por campo da fonte: não")
            if parameter.get("used_in_dashboards"):
                lines.append("- Usado nos painéis:")
                for dashboard in sorted(parameter["used_in_dashboards"], key=str.lower):
                    lines.append(f"  - {dashboard}")
            else:
                lines.append("- Usado nos painéis: não está sendo usado em nenhum painel")
            if parameter.get("used_in_worksheets"):
                lines.append("- Usado nas planilhas:")
                for worksheet in sorted(parameter["used_in_worksheets"], key=str.lower):
                    lines.append(f"  - {worksheet}")
            else:
                lines.append("- Usado nas planilhas: não está sendo usado em nenhuma planilha")
            lines.append("")
        return lines

    def _build_calculations_markdown(self) -> list[str]:
        lines = ["## Campos Calculados", ""]
        calculations = self.metadata["calculations"]
        if not calculations:
            lines.append("- Nenhum campo calculado encontrado.")
            return lines

        grouped: dict[str, list[dict[str, Any]]] = {}
        for calculation in calculations:
            grouped.setdefault(calculation["datasource"] or "(sem datasource)", []).append(calculation)

        for datasource, items in grouped.items():
            lines.append(f"### {datasource}")
            lines.append("")
            for item in items:
                lines.append(f"- Campo: `{item.get('caption')}`")
                lines.append(f"- Nome interno: `{item.get('name')}`")
                lines.append(f"- Origem: `{item.get('origin')}`")
                codigo = item.get("codigo") or "-"
                lines.append(f"- Código: {codigo}")
                impacts = sorted(item.get("impacts", []), key=str.lower)
                if impacts:
                    lines.append("- Impacta / é referenciado por:")
                    for impact in impacts:
                        lines.append(f"  - {impact}")
                else:
                    lines.append("- Impacta / é referenciado por: nenhum outro campo calculado")
                if item.get("used_in_dashboards"):
                    lines.append("- Usado nos painéis:")
                    for dashboard in sorted(item["used_in_dashboards"], key=str.lower):
                        lines.append(f"  - {dashboard}")
                else:
                    lines.append("- Usado nos painéis: não está sendo usado em nenhum painel")
                if item.get("used_in_worksheets"):
                    lines.append("- Usado nas planilhas:")
                    for worksheet in sorted(item["used_in_worksheets"], key=str.lower):
                        lines.append(f"  - {worksheet}")
                else:
                    lines.append("- Usado nas planilhas: não está sendo usado em nenhuma planilha")
                lines.append("")
        return lines

    def _build_dashboards_markdown(self) -> list[str]:
        lines = ["## Dashboards", ""]
        dashboards = self.metadata["dashboards"]
        if not dashboards:
            lines.append("- Nenhum dashboard encontrado.")
            return lines

        for dashboard in dashboards:
            lines.append(f"### {dashboard['name']}")
            lines.append("")
            if dashboard.get("worksheet_members"):
                lines.append("- Worksheets incluídas:")
                for member in sorted(dashboard["worksheet_members"], key=str.lower):
                    worksheet = next(
                        (item for item in self.metadata["worksheets"] if item.get("name") == member),
                        None,
                    )
                    datasources = worksheet.get("datasource_dependencies", []) if worksheet else []
                    lines.append(f"  - {member}")
                    if datasources:
                        lines.append(f"    Fonte de dados: {', '.join(sorted(datasources, key=str.lower))}")
                    else:
                        lines.append("    Fonte de dados: nenhuma identificada")
            else:
                lines.append("- Worksheets incluídas: nenhuma identificada")
            if dashboard.get("filters_used"):
                exposed_filters = [item for item in dashboard["filters_used"] if item.get("visibility") == "exposto no painel"]
                internal_filters = [item for item in dashboard["filters_used"] if item.get("visibility") == "interno da planilha"]
                if exposed_filters:
                    lines.append("- Filtros expostos no painel:")
                    for filter_item in exposed_filters:
                        context_label = "sim" if filter_item.get("is_context") else "não"
                        field_label = filter_item.get("field_label") or "(sem nome)"
                        filter_class = filter_item.get("class") or "-"
                        lines.append(f"  - {field_label} | tipo: {filter_class} | contexto: {context_label}")
                else:
                    lines.append("- Filtros expostos no painel: nenhum identificado")
                if internal_filters:
                    lines.append("- Filtros internos das planilhas:")
                    for filter_item in internal_filters:
                        context_label = "sim" if filter_item.get("is_context") else "não"
                        field_label = filter_item.get("field_label") or "(sem nome)"
                        filter_class = filter_item.get("class") or "-"
                        lines.append(f"  - {field_label} | tipo: {filter_class} | contexto: {context_label}")
                else:
                    lines.append("- Filtros internos das planilhas: nenhum identificado")
            else:
                lines.append("- Filtros expostos no painel: nenhum identificado")
                lines.append("- Filtros internos das planilhas: nenhum identificado")
            lines.append(f"- Zonas no layout: {len(dashboard.get('zones', []))}")
            lines.append(f"- Layouts de device: {len(dashboard.get('device_layouts', []))}")
            unique_colors = sorted(set(dashboard.get("colors_used", [])), key=str.lower)
            unique_fonts = sorted(set(dashboard.get("fonts_used", [])), key=str.lower)
            lines.append(f"- Cores usadas: {', '.join(unique_colors) or '-'}")
            lines.append(f"- Fontes usadas: {', '.join(unique_fonts) or '-'}")
            lines.append("")
        return lines

    def _build_visual_tokens_markdown(self) -> list[str]:
        lines = ["## Tokens Visuais", ""]
        colors = self.metadata["visual_tokens"]["colors"]
        fonts = self.metadata["visual_tokens"]["fonts"]
        unique_color_values = sorted({item["value"] for item in colors if item.get("value")}, key=str.lower)
        unique_font_values = sorted({item["value"] for item in fonts if item.get("value")}, key=str.lower)

        lines.append(f"- Total de cores únicas: {len(unique_color_values)}")
        lines.append(f"- Total de fontes únicas: {len(unique_font_values)}")
        lines.append("")
        lines.append("### Cores")
        lines.append("")
        if unique_color_values:
            for color in unique_color_values:
                lines.append(f"- `{color}`")
        else:
            lines.append("- Nenhuma cor identificada.")
        lines.append("")
        lines.append("### Fontes")
        lines.append("")
        if unique_font_values:
            for font in unique_font_values:
                lines.append(f"- `{font}`")
        else:
            lines.append("- Nenhuma fonte identificada.")
        lines.append("")
        return lines

    def _build_hyper_markdown(self) -> list[str]:
        lines = ["## Estrutura dos Extracts Hyper", ""]
        hyper_extracts = self.metadata.get("hyper_extracts", [])
        if not hyper_extracts:
            lines.append("- Nenhum extract Hyper identificado no workbook.")
            return lines

        lines.append("- Observação: a estrutura abaixo foi inferida a partir do metadata do workbook; os arquivos `.hyper` não foram copiados para `data`.")
        lines.append("")
        for extract in hyper_extracts:
            lines.append(f"### {extract.get('datasource')}")
            lines.append("")
            lines.append(f"- Arquivo Hyper relacionado: `{extract.get('hyper_file') or '-'}`")
            lines.append(f"- Quantidade de campos: {extract.get('field_count', 0)}")
            lines.append("- Campos:")
            for field in extract.get("fields", []):
                lines.append(
                    f"  - {field['field']} | datatype: {field.get('datatype') or '-'} | formato: {field.get('format') or '-'} | papel: {field.get('role') or '-'}"
                )
            lines.append("")
        return lines

    def _build_package_manifest_markdown(self) -> list[str]:
        lines = ["## Conteúdo do Pacote", ""]
        if not self.package_manifest:
            lines.append("- Nenhum membro de pacote registrado.")
            return lines
        for item in self.package_manifest:
            lines.append(f"- `{item['path']}` ({item['kind']}, {item['size_bytes']} bytes)")
        return lines

    def _write_excel(self) -> Path:
        output_path = self.output_dir / f"{self.base_name}.xlsx"
        with pd.ExcelWriter(output_path) as writer:
            self._to_frame(self.metadata["parameters"]).to_excel(writer, sheet_name="Parameters", index=False)
            self._to_frame(self.metadata["calculations"]).to_excel(writer, sheet_name="Calculations", index=False)
            self._to_frame(self.metadata["datasources"], stringify_nested=True).to_excel(writer, sheet_name="Datasources", index=False)
            self._to_frame(self.metadata["worksheets"], stringify_nested=True).to_excel(writer, sheet_name="Worksheets", index=False)
            self._to_frame(self.metadata["dashboards"], stringify_nested=True).to_excel(writer, sheet_name="Dashboards", index=False)
            self._to_frame(self.metadata["visual_tokens"]["colors"]).to_excel(writer, sheet_name="Colors", index=False)
            self._to_frame(self.metadata["visual_tokens"]["fonts"]).to_excel(writer, sheet_name="Fonts", index=False)
            self._to_frame(self.package_manifest).to_excel(writer, sheet_name="PackageManifest", index=False)
        return output_path

    def _to_frame(self, records: list[dict[str, Any]], stringify_nested: bool = False) -> pd.DataFrame:
        if not records:
            return pd.DataFrame()
        df = pd.DataFrame(records)
        if stringify_nested:
            for column in df.columns:
                df[column] = df[column].apply(
                    lambda value: json.dumps(value, ensure_ascii=False, indent=2)
                    if isinstance(value, (dict, list))
                    else value
                )
        return df


def load_path_from_config(config_path: str | Path | None = None) -> str:
    """Lê o caminho padrão do arquivo Tableau a partir do config.json."""
    config_file = Path(config_path) if config_path is not None else PROJECT_ROOT / "config" / "config.json"
    if not config_file.exists():
        raise FileNotFoundError(
            f"Nenhum caminho informado e o arquivo '{config_file}' não foi encontrado."
        )

    try:
        config_data = json.loads(config_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"O arquivo '{config_file}' não contém um JSON válido.") from exc

    tableau_path = config_data.get("tableau_path")
    if not tableau_path:
        raise ValueError(
            f"O arquivo '{config_file}' não possui a chave obrigatória 'tableau_path'."
        )
    return tableau_path


def parse_args() -> argparse.Namespace:
    """Define a interface de linha de comando."""
    parser = argparse.ArgumentParser(
        description="Gera documentação de workbooks Tableau (.twb/.twbx)."
    )
    parser.add_argument(
        "filepath",
        nargs="?",
        help="Caminho completo do arquivo Tableau (.twb ou .twbx).",
    )
    parser.add_argument(
        "--format",
        dest="output_format",
        choices=["all", "markdown", "json", "excel"],
        default="all",
        help="Formato principal da documentação a gerar. O mapa XPath/JSON é sempre gerado.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        file_path = args.filepath or load_path_from_config()
        documenter = TableauDoc(file_path, output_format=args.output_format)
        written_files = documenter.write_outputs()

        print(f"Arquivo fonte: {documenter.source_path}")
        print(f"Diretório de saída: {documenter.output_dir}")
        print("Arquivos gerados:")
        for path in written_files:
            print(f"- {path}")
    except (FileNotFoundError, ValueError, ET.ParseError, zipfile.BadZipFile) as exc:
        print(f"Erro: {exc}")
        print("Uso: python Tableau_doc.py <caminho_do_arquivo.twb|.twbx> [--format all|markdown|json|excel]")
        sys.exit(1)


if __name__ == "__main__":
    main()
