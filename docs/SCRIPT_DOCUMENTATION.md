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
- `all`

## Configuração

Quando o caminho do workbook não é informado na linha de comando, o script tenta carregar o valor padrão a partir de:

```text
config/config.json
```

## Fluxo geral

1. valida o caminho do arquivo de entrada
2. detecta se o arquivo é ` .twb` ou ` .twbx`
3. em caso de ` .twbx`, extrai o conteúdo do pacote para `package_contents/`
4. lê o workbook XML principal
5. monta estruturas internas de metadados
6. gera o mapa XPath/JSON
7. gera a documentação no formato solicitado
8. remove artefatos temporários de processamento

## Estruturas extraídas

### Workbook

- versão do Tableau
- build de origem
- plataforma de origem
- caminho do arquivo
- data/hora da última alteração do arquivo
- `repository-location`, quando disponível

### Fontes de dados

- nome interno
- caption
- versão
- indicação de fonte publicada
- metadados de repositório publicado, quando presentes
- colunas e metadados associados

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
- código
- dependência de outros cálculos
- uso em painéis e planilhas
- indicação explícita quando não houver uso

### Tokens visuais

- cores únicas
- fontes únicas
- paletas de cor

## Decisões importantes de implementação

### Diretório de saída fixo

O script não usa o diretório atual de execução para gravar resultados.  
Ele sempre grava em `data/` na raiz do projeto.

### ` .hyper`

Os arquivos ` .hyper` não são copiados para a pasta `data/`.

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

## Limitações conhecidas

- nem todo workbook publicado carrega no XML a estrutura completa da fonte publicada
- a detecção de filtro de contexto depende do que estiver explícito no XML
- alguns nomes técnicos de campos podem permanecer se o workbook não trouxer caption associado
- a estrutura dos extracts remotos não é introspectada binariamente

## Dependências

- `pandas`
- `openpyxl`

## Manutenção sugerida

Se o script evoluir, priorizar:

1. ampliar a tradução de nomes técnicos remanescentes
2. melhorar a classificação de filtros de contexto
3. ampliar a detecção de uso indireto de parâmetros e cálculos
4. adicionar testes automatizados para workbooks de exemplo
