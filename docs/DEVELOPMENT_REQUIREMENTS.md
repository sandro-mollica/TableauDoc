# Requisitos Usados no Desenvolvimento

Este documento consolida os requisitos funcionais, técnicos e de saída aplicados ao desenvolvimento do `Tableau_doc.py`.

## Requisitos funcionais

1. Ler arquivos Tableau ` .twb` e ` .twbx`
2. Extrair o workbook XML principal
3. Gravar o conteúdo processado em `data/<nome-do-arquivo>/`
4. Usar sempre o nome-base do arquivo de entrada nos artefatos de saída
5. Gerar documentação em:
   - Markdown
   - JSON
   - Excel
6. Gerar um documento `mapa_XPath_JSON`
7. Extrair miniaturas embutidas quando existirem
8. Documentar:
   - fontes de dados
   - dashboards
   - worksheets
   - parâmetros
   - campos calculados
   - filtros
   - cores
   - fontes
   - paletas

## Requisitos de saída

1. A saída deve ser gravada sempre em `data/`, na raiz do projeto
2. O script não deve depender do diretório atual de execução
3. Com `--format markdown`, não deve gerar `json` nem `xlsx`
4. Com `--format json`, não deve gerar `md` nem `xlsx`
5. Com `--format excel`, não deve gerar `md` nem `json`
6. Com `--format all`, deve gerar todos os formatos
7. Após a execução, o script deve remover artefatos temporários não necessários ao resultado final

## Requisitos de conteúdo do relatório

### Resumo

1. Exibir caminho do arquivo
2. Exibir tipo do arquivo
3. Exibir diretório de saída
4. Exibir quantidade de datasources
5. Exibir quantidade de parâmetros
6. Exibir quantidade de campos calculados
7. Exibir quantidade de worksheets
8. Exibir quantidade de dashboards
9. Exibir data e hora da última alteração do arquivo

### Fontes de dados

1. Listar todas as fontes de dados usadas
2. Indicar quando a fonte for publicada
3. Mostrar metadados de repositório quando presentes

### Dashboards

1. Mostrar dashboards antes de parâmetros e cálculos
2. Listar as worksheets que compõem cada dashboard
3. Em cada worksheet listada, mostrar a fonte de dados usada
4. Separar filtros em:
   - filtros expostos no painel
   - filtros internos das planilhas
5. Indicar contexto quando detectado

### Parâmetros

1. Mostrar nome interno e nome real
2. Mostrar tipo e domínio
3. Mostrar valor atual/default
4. Mostrar membros de lista quando existirem
5. Mostrar quando um parâmetro for preenchido por um campo da fonte
6. Mostrar em quais painéis o parâmetro é usado
7. Mostrar em quais planilhas o parâmetro é usado
8. Informar explicitamente quando não estiver sendo usado

### Campos calculados

1. Mostrar nome interno e nome real
2. Mostrar origem
3. Mostrar o código como item `Código`
4. Preservar comentários do cálculo
5. Não usar bloco fenced code para o item `Código`
6. Mostrar uso em painéis
7. Mostrar uso em planilhas
8. Informar explicitamente quando não houver uso
9. Mostrar `Impacta / é referenciado por` em lista vertical
10. Ordenar a lista alfabeticamente
11. Não repetir campos na lista
12. Exibir apenas nome real, não nome interno

### Tokens visuais

1. Não repetir cores
2. Não repetir fontes
3. Ignorar falsos positivos como `match` e `Vertical`
4. Mostrar paletas de cor separadamente

## Requisitos específicos sobre ` .hyper`

1. O script não deve copiar arquivos ` .hyper` para `data/`
2. O item `Estrutura dos Extracts Hyper` foi removido do relatório final

## Requisitos de limpeza

1. O diretório `package_contents/` deve ser removido ao final da execução
2. Apenas os artefatos finais devem permanecer no diretório de saída
3. A rotina de limpeza deve ser extensível para futuros artefatos temporários

## Requisitos de organização do relatório

Ordem definida para o arquivo de saída:

1. Título
2. Resumo
3. Workbook
4. Repository Location
5. Fontes de Dados
6. Dashboards
7. Tokens Visuais
8. Preferências e Paletas
9. Paletas de Cor
10. Cores
11. Fontes
12. Parâmetros
13. Campos Calculados

## Requisitos técnicos

1. Implementação em Python
2. Uso de `pandas` para a saída Excel
3. Uso de `openpyxl` via `pandas.ExcelWriter`
4. Uso de `xml.etree.ElementTree` para leitura do XML
5. Compatibilidade com caminhos absolutos locais
6. Compatibilidade com nomes de arquivo contendo espaços e acentos
7. O arquivo de configuração padrão deve ficar em `config/config.json`

## Requisitos de manutenção

1. O projeto deve ter um `.gitignore`
2. O `.gitignore` deve ignorar `data/`
3. O `.gitignore` deve ignorar `img/`
4. O `requirements.txt` deve refletir apenas dependências realmente usadas
