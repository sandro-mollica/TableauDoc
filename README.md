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
python3 Tableau_doc.py /caminho/arquivo.twbx --format rtf
```

## Leitura opcional de `.tdsx` externo

Quando o SQL customizado não estiver dentro do `.twb/.twbx`, o script pode procurar um `.tdsx` externo.

Essa busca só acontece quando a chave `external_tdsx_paths` existir no `config/config.json`.

O script processa um único arquivo `.twb` ou `.twbx` por execução.
Esse arquivo, porém, pode ser associado a um ou mais `.tdsx` externos.

Para `external_tdsx_paths`, o script aceita dois formatos:

- `string`: quando houver apenas um `.tdsx` externo
- `lista`: quando houver mais de um `.tdsx` ou quando você quiser misturar arquivos e pastas

Exemplo com um único `.tdsx`:

```json
{
  "tableau_path": "/caminho/arquivo.twbx",
  "external_tdsx_paths": "/caminho/fonte_publicada.tdsx"
}
```

Exemplo de `config/config.json`:

```json
{
  "tableau_path": "/caminho/arquivo.twbx",
  "external_tdsx_paths": [
    "/caminho/pasta_com_tdsx",
    "/caminho/fonte_publicada.tdsx"
  ]
}
```

Cada item de `external_tdsx_paths` pode ser:

- uma pasta, para busca recursiva de arquivos `.tdsx`
- um arquivo `.tdsx` específico

Quando houver mais de um `.tdsx`, o código percorre todos os caminhos informados e tenta associar cada datasource do workbook aos arquivos externos compatíveis pelo nome.

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

- `--format markdown`: gera `xml`, `mapa_XPath_JSON.*`, `md` e `rtf`
- `--format json`: gera `xml`, `mapa_XPath_JSON.*` e `json`
- `--format excel`: gera `xml`, `mapa_XPath_JSON.*` e `xlsx`
- `--format rtf`: gera `xml`, `mapa_XPath_JSON.*` e `rtf`
- `--format all`: gera todos os formatos

## O que o script documenta

- fontes de dados
- tipo da fonte (`lógico` ou `físico (join)`)
- modo de conexão (`Live`, `Extração`, `Live + extração habilitada`, `Live + extração desabilitada`)
- arquivos `.hyper` associados, quando existirem
- campos da fonte de dados, incluindo nome interno, datatype, role, alias e indicação de oculto
- mapa de relacionamentos entre tabelas
- dashboards e planilhas
- parâmetros
- campos calculados e dependências
- indicação de uso ou não uso de objetos
- filtros expostos e filtros internos
- cores, fontes e paletas
- mapa XPath/JSON do workbook
- miniaturas embutidas

## Destaques do relatório

- o arquivo `.rtf` usa `Calibre` como fonte principal e uma fonte monoespaçada padrão do sistema para trechos de código e valores hexadecimais
- o relatório mostra data e hora de geração logo abaixo do título
- a seção `Relação de Campos Calculados` informa alias, uso, dependências, código e se o campo está oculto
- a seção `Objetos não usados` lista campos calculados, parâmetros e fontes de dados sem uso efetivo
- quando o XML não explicita a tipografia usada em um dashboard, o relatório mostra `Fonte padrão do Tableau (não explicitada no XML)`

## Observações

- arquivos ` .hyper` não são copiados para `data/`
- artefatos temporários de processamento são removidos automaticamente ao final da execução
- a data de última alteração do workbook é exibida no resumo do relatório
