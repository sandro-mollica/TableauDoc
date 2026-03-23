# TableauDoc

Ferramenta em Python para documentar workbooks Tableau ` .twb` e ` .twbx` sem abrir o Tableau Desktop.

## Arquivos principais

- [Tableau_doc.py](/Users/sandromollica/Library/CloudStorage/OneDrive-Pessoal/Workspaces/Antigravity/TableauDoc/Tableau_doc.py): launcher de compatibilidade na raiz
- [src/tableau_doc.py](/Users/sandromollica/Library/CloudStorage/OneDrive-Pessoal/Workspaces/Antigravity/TableauDoc/src/tableau_doc.py): implementação principal
- [requirements.txt](/Users/sandromollica/Library/CloudStorage/OneDrive-Pessoal/Workspaces/Antigravity/TableauDoc/requirements.txt): dependências Python
- [docs/SCRIPT_DOCUMENTATION.md](/Users/sandromollica/Library/CloudStorage/OneDrive-Pessoal/Workspaces/Antigravity/TableauDoc/docs/SCRIPT_DOCUMENTATION.md): documentação funcional e técnica do script
- [docs/DEVELOPMENT_REQUIREMENTS.md](/Users/sandromollica/Library/CloudStorage/OneDrive-Pessoal/Workspaces/Antigravity/TableauDoc/docs/DEVELOPMENT_REQUIREMENTS.md): requisitos usados no desenvolvimento

## Instalação

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
python3 Tableau_doc.py /caminho/arquivo.twbx --format all
python3 Tableau_doc.py /caminho/arquivo.twbx --format markdown
python3 Tableau_doc.py /caminho/arquivo.twbx --format json
python3 Tableau_doc.py /caminho/arquivo.twbx --format excel
```

## Saída

Os artefatos são sempre gravados em:

```bash
data/<nome-do-arquivo>/
```

Esse diretório é criado na raiz do projeto, independentemente do diretório atual de execução.

## Estrutura sugerida

```text
TableauDoc/
├─ config/
│  └─ config.json
├─ src/
│  ├─ __init__.py
│  └─ tableau_doc.py
├─ docs/
│  ├─ SCRIPT_DOCUMENTATION.md
│  └─ DEVELOPMENT_REQUIREMENTS.md
├─ data/
├─ TableauDoc.code-workspace
├─ Tableau_doc.py
├─ README.md
├─ requirements.txt
└─ .gitignore
```

## Formatos

- `--format markdown`: gera `xml`, `mapa_XPath_JSON.*` e `md`
- `--format json`: gera `xml`, `mapa_XPath_JSON.*` e `json`
- `--format excel`: gera `xml`, `mapa_XPath_JSON.*` e `xlsx`
- `--format all`: gera todos os formatos

## O que o script documenta

- fontes de dados
- dashboards e planilhas
- parâmetros
- campos calculados e dependências
- filtros expostos e filtros internos
- cores, fontes e paletas
- mapa XPath/JSON do workbook
- miniaturas embutidas

## Observações

- arquivos ` .hyper` não são copiados para `data/`
- artefatos temporários de processamento são removidos automaticamente ao final da execução
- a data de última alteração do workbook é exibida no resumo do relatório
