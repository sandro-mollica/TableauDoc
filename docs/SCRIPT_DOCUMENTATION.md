# Documentação do Script `Tableau_doc.py`

## Objetivo

O launcher `Tableau_doc.py` na raiz encaminha a execução para a implementação principal em `src/tableau_doc.py`, responsável por extrair e documentar metadados de arquivos Tableau ` .twb` e ` .twbx` sem depender do Tableau Desktop.

Ele transforma o workbook em artefatos legíveis para análise, auditoria e documentação técnica.

## Entradas suportadas

- ` .twb`
- ` .twbx`

## Saídas geradas

Os arquivos são gravados sempre em:

```text
<raiz-do-projeto>/data/<nome-base-do-arquivo>/
```

### Artefatos possíveis

- `<arquivo>.xml`
- `mapa_XPath_JSON.md`
- `mapa_XPath_JSON.json`
- `<arquivo>.md`
- `<arquivo>.rtf`
- `<arquivo>.json`
- `<arquivo>.xlsx`
- `<arquivo>_manifest.json`
- `thumbnails/`

## Parâmetros de execução

```bash
python3 Tableau_doc.py /caminho/arquivo.twbx --format all
```

### Opções de `--format`

- `markdown`
- `json`
- `excel`
- `rtf`
- `all`

## Configuração

Quando o caminho do workbook não é informado na linha de comando, o script tenta carregar o valor padrão a partir de:

```text
config/config.json
```

Também é possível definir caminhos opcionais para leitura de `.tdsx` externo:

Formato aceito para `tableau_path`:

- `string`, para um único arquivo Tableau

O script processa um único arquivo `.twb` ou `.twbx` por execução.
Esse arquivo pode ter um ou mais `.tdsx` externos associados.

Formato aceito para `external_tdsx_paths`:

- `string`, para um único `.tdsx`
- `lista`, para múltiplos `.tdsx` e/ou pastas

Exemplo com um único `.tdsx`:

```json
{
  "tableau_path": "/caminho/arquivo.twbx",
  "external_tdsx_paths": "/caminho/fonte_publicada.tdsx"
}
```

Exemplo com múltiplos caminhos:

```json
{
  "tableau_path": "/caminho/arquivo.twbx",
  "external_tdsx_paths": [
    "/caminho/pasta_com_tdsx",
    "/caminho/fonte_publicada.tdsx"
  ]
}
```

Condição aplicada no código:

- se `external_tdsx_paths` não estiver definido, o script não tenta ler `.tdsx` externo
- se `external_tdsx_paths` existir, a busca externa é habilitada

Cada item pode apontar para:

- uma pasta com busca recursiva por arquivos `.tdsx`
- um arquivo `.tdsx` específico

Se houver mais de um `.tdsx`, o script percorre todos os caminhos informados e tenta localizar correspondências por datasource para extrair o SQL customizado disponível.

## Fluxo geral

1. valida o caminho do arquivo de entrada
2. detecta se o arquivo é ` .twb` ou ` .twbx`
3. em caso de ` .twbx`, extrai o conteúdo do pacote para `package_contents/`
4. lê o workbook XML principal
5. monta estruturas internas de metadados
6. gera o mapa XPath/JSON
7. gera a documentação no formato solicitado
8. quando o formato é `markdown`, gera em conjunto os arquivos `md` e `rtf`
9. remove artefatos temporários de processamento

## Estruturas extraídas

### Workbook

- versão do Tableau
- build de origem
- plataforma de origem
- caminho do arquivo
- data/hora da última alteração do arquivo
- `repository-location`, quando disponível

### Fontes de dados

- caption
- tipo da fonte: `lógico` ou `físico (join)`
- versão
- indicação de fonte publicada
- modo de conexão por conexão
- caminhos `.hyper` associados, quando existirem
- metadados de repositório publicado, quando presentes
- SQL customizado embutido no `.twb/.twbx`, quando existir
- SQL customizado encontrado em `.tdsx` externo, quando a busca opcional estiver habilitada
- mapa de relacionamentos entre tabelas
- colunas e metadados associados
- aliases dos campos, quando existirem
- indicação se o campo está oculto ou não

### Dashboards

- nome do dashboard
- worksheets incluídas
- fonte de dados usada por cada worksheet
- filtros expostos no painel
- filtros internos das planilhas
- layout e zonas
- cores e fontes usadas

### Parâmetros

- nome interno
- caption
- tipo
- domínio
- valor atual/default
- membros de lista
- origem `source-field`, quando houver
- uso em painéis e planilhas
- indicação explícita quando não houver uso

### Campos calculados

- nome interno
- nome real
- origem
- aliases, quando existirem
- indicação se o campo está oculto ou não
- dependência de outros cálculos
- uso em painéis e planilhas
- código exibido por último na ficha de cada campo calculado
- indicação explícita quando não houver uso

### Objetos não usados

- campos calculados sem uso em painéis, planilhas ou outros objetos
- parâmetros sem uso em painéis, planilhas ou outros objetos
- fontes de dados sem uso efetivo
- nome real e aliases, quando existirem

### Tokens visuais

- cores únicas
- fontes únicas
- paletas de cor

### Diretório de saída fixo

O script não usa o diretório atual de execução para gravar resultados.  
Ele sempre grava em `data/` na raiz do projeto.

### Limpeza pós-processamento

Após a geração dos artefatos finais, o script remove diretórios e arquivos temporários de trabalho registrados internamente na rotina de limpeza, como:

- `package_contents/`
- `tmp/`
- `temp/`
- `.tmp/`

### Dedupliação

O script deduplica:

- cores
- fontes
- dependências entre campos calculados
- listas compostas reaproveitadas em dashboards e planilhas

### Humanização de nomes

O script tenta substituir nomes técnicos por captions reais sempre que encontra mapeamento suficiente no XML.

Isso é usado em:

- filtros
- campos calculados
- referências entre cálculos
- datasources dependentes nas planilhas

### Tipografia do RTF

O relatório `.rtf` usa:

- `Calibre` como fonte principal
- uma fonte monoespaçada padrão do sistema para código de campos calculados e valores hexadecimais

Os nomes das fontes ficam explícitos no código para facilitar personalização.

## Limitações conhecidas

- nem todo workbook publicado carrega no XML a estrutura completa da fonte publicada
- nem todo `.twbx` embute o `.tdsx`; em muitos casos o SQL customizado só existe fora do pacote
- a detecção de filtro de contexto depende do que estiver explícito no XML
- alguns nomes técnicos de campos podem permanecer se o workbook não trouxer caption associado

## Dependências

- `pandas`
- `openpyxl`
