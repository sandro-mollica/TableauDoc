# Referência de Rotinas do `tableaudoc/tableau_doc.py`

## Visão geral

Este documento descreve, de forma sistemática, o papel de cada rotina presente em `tableaudoc/tableau_doc.py`.

O objetivo é oferecer uma leitura técnica orientada à manutenção. A proposta não é apenas listar nomes de funções, mas explicar como cada rotina participa do fluxo de carga do workbook, extração de metadados, enriquecimento semântico, geração de artefatos e integração com a linha de comando.

## Convenções de leitura

- Funções de módulo: utilitários independentes, usados antes ou fora da classe principal.
- Métodos da classe `TableauDoc`: rotinas de processamento vinculadas ao estado do workbook carregado.
- Rotinas de saída: responsáveis por converter metadados em arquivos finais.
- Rotinas de CLI: responsáveis por configuração, argumentos e execução principal.

## Funções utilitárias de módulo

### `sanitize_filename`

Normaliza um texto para uso seguro como nome de arquivo. Remove caracteres problemáticos, preserva legibilidade e garante um fallback estável quando o valor de entrada é vazio.

### `decode_tableau_text`

Converte entidades XML recorrentes do Tableau para texto legível. É especialmente útil para fórmulas e trechos de SQL armazenados com escapes como `&quot;`, `&lt;` e quebras de linha numéricas.

### `clean_brackets`

Remove colchetes de identificadores internos do Tableau. A rotina é usada para transformar nomes técnicos como `[Campo]` em uma forma mais adequada para exibição e comparação.

### `clean_display_label`

Faz limpeza visual em labels extraídos do XML. Compacta espaços, remove ruídos de formatação e melhora a apresentação de textos exibidos em relatórios.

### `is_color_like`

Avalia se um valor se comporta como cor ou paleta. Reconhece códigos hexadecimais, funções como `rgb(...)` e nomes conhecidos de cores.

### `unique_ordered`

Remove duplicados sem perder a ordem de aparição. A rotina suporta listas com valores simples e também estruturas aninhadas como listas e dicionários.

### `compact_json`

Serializa estruturas aninhadas em uma única linha legível. É usada principalmente quando o relatório precisa exibir atributos complexos sem expandir um bloco inteiro de JSON.

### `format_yes_no`

Traduz indicadores booleanos ou equivalentes textuais para `sim` e `não`. Essa função padroniza a linguagem de saída dos relatórios.

### `_log_progress`

Centraliza a emissão de mensagens de progresso no console. A rotina existe para tornar a execução observável sem espalhar `print` diretamente pelo código.

### `normalize_whitespace`

Normaliza quebras de linha e remove excesso de espaços nas extremidades. É aplicada em fórmulas, SQL e textos livres extraídos do XML.

### `normalize_lookup_token`

Converte nomes em tokens comparáveis de forma flexível. O foco é facilitar associação entre nomes de datasources, arquivos externos e metadados publicados.

### `element_path_with_indices`

Calcula um XPath legível para um elemento XML, inclusive quando há irmãos com a mesma tag.

#### `walk`

Função interna usada exclusivamente por `element_path_with_indices`. Percorre recursivamente a árvore XML até encontrar o nó de destino e montar o caminho correspondente.

## Classe `TableauDoc`

### `__init__`

Inicializa a instância principal do processamento. A rotina valida o arquivo de entrada, prepara diretório de saída, registra parâmetros de execução, carrega o XML do workbook e constrói a estrutura consolidada de metadados.

### `_load_root_and_extract_contents`

Carrega o XML principal do workbook. Se a origem for `.twb`, copia o arquivo para a área de trabalho; se for `.twbx`, extrai o pacote, preservando os itens necessários para análise e ignorando arquivos `.hyper`.

### `_classify_package_member`

Classifica cada item do pacote em categorias simples, como workbook, imagem, extração ou arquivo de apoio. Essa classificação é usada no manifesto final.

### `_build_metadata`

Orquestra a extração completa de informações do workbook. A rotina chama os extratores especializados, aplica enriquecimentos e devolve a estrutura-mestre consumida por JSON, Markdown, RTF, DOCX e Excel.

### `_build_caption_lookup`

Monta dicionários de tradução entre nomes técnicos e captions amigáveis. Essa camada é essencial para substituir referências internas por nomes mais compreensíveis nos relatórios.

### `_extract_repository_location`

Lê a tag `repository-location`, quando disponível, e a converte em dicionário. A rotina é usada tanto no workbook quanto em datasources publicados.

### `_extract_preferences`

Extrai preferências globais e paletas de cor definidas no workbook. O resultado compõe a seção de preferências e tokens visuais.

### `_extract_global_styles`

Extrai regras globais de estilo a partir da estrutura `style/style-rule`. É a base para observar tipografia e formatação declaradas em nível de workbook.

### `_extract_datasources`

Percorre todas as fontes de dados e consolida conexões, campos, metadados, cálculos de conexão, relacionamentos e blocos de extração. É uma das rotinas centrais do pipeline.

### `_extract_extracts`

Lê blocos `extract` de um datasource e transforma seus atributos em estruturas Python. Serve de apoio para inferência de modo de conexão e caminhos `.hyper`.

### `_infer_connection_mode`

Classifica uma conexão como `Live`, `Extração`, `Live + extração habilitada` ou `Live + extração desabilitada`. A lógica combina atributos da conexão e presença de blocos de extração.

### `_extract_hyper_paths`

Isola caminhos de arquivos `.hyper` associados a extrações. O resultado é deduplicado e reaproveitado nas saídas documentais.

### `_enrich_datasources_with_external_tdsx`

Complementa datasources com informações externas vindas de `.tdsx`, quando essa busca opcional está habilitada. Anexa SQL customizado e mapas de relacionamento externos.

### `_find_matching_external_tdsx`

Localiza arquivos `.tdsx` candidatos para um datasource específico. A rotina compara captions, nomes internos e referências de repositório com nomes de arquivo e caminhos disponíveis.

### `_extract_custom_sql_from_tdsx`

Abre um `.tdsx`, localiza arquivos `.tds` internos e tenta extrair blocos de SQL customizado. O foco é recuperar consultas que não estejam presentes no workbook principal.

### `_extract_relationship_map_from_tdsx`

Lê mapas de relacionamento a partir de um `.tdsx`. O método reutiliza a rotina genérica de extração de relacionamentos a partir de uma raiz XML.

### `_extract_relationship_map_from_root`

Extrai relações entre objetos a partir de qualquer bloco XML que contenha `object-graph` e `relationships`. É usado tanto no workbook quanto em fontes externas.

### `_humanize_relationship_endpoint_attributes`

Substitui `object-id` técnicos por rótulos mais legíveis nas extremidades de um relacionamento. O objetivo é melhorar a interpretação do mapa gerado.

### `_relationship_expression_to_text`

Converte a expressão XML de um relacionamento em texto legível. Essa saída é usada para mostrar a condição de ligação entre tabelas.

### `_relationship_link_fields`

Tenta identificar os campos de origem e destino usados em um relacionamento. O retorno é útil para diagramas textuais e resumos de join.

### `_parse_connection`

Transforma uma tag `connection` em uma estrutura Python com atributos, relações e conexões nomeadas. É a base para documentar o comportamento técnico da origem de dados.

### `_extract_relation_sql`

Detecta SQL customizado em uma `relation`. A rotina avalia atributos, texto agregado e indícios semânticos para decidir se o conteúdo parece uma consulta SQL.

### `_parse_column`

Converte uma coluna de datasource em dicionário estruturado. Inclui atributos básicos, aliases, membros, bins e fórmulas associadas.

### `_parse_metadata_record`

Extrai um `metadata-record` com seus atributos e nomes locais/remotos. Essa informação ajuda a documentar campos oriundos de fontes físicas ou publicadas.

### `_extract_parameters`

Localiza e devolve os parâmetros do workbook, tratados como colunas do datasource especial `Parameters`.

### `_extract_calculations`

Constrói a lista consolidada de campos calculados a partir de colunas, cálculos de conexão e metadata records. Também prepara nomes amigáveis, impacto entre cálculos e versões de fórmula mais legíveis.

### `_replace_internal_names`

Substitui nomes internos em fórmulas por captions mais amigáveis. A rotina melhora a legibilidade dos cálculos no relatório final.

### `_enrich_usage`

Analisa uso de parâmetros e cálculos em planilhas e dashboards. O resultado marca onde cada objeto aparece e sustenta as seções de uso e não uso.

### `_enrich_datasource_field_usage`

Complementa o uso em nível de campo dentro de datasources. A rotina cruza referências de planilhas, cálculos e parâmetros para identificar utilização efetiva.

### `_xml_contains_reference`

Verifica se um texto XML contém referência a um nome interno ou caption. É um utilitário de busca semântica usado em rotinas de enriquecimento.

### `_parse_source_field_reference`

Interpreta o valor de `source-field` e o transforma em estrutura legível. Essa informação é usada principalmente em parâmetros ligados a campos da fonte.

### `_humanize_datasource_reference`

Traduz referências técnicas de datasource para captions conhecidas. É uma etapa importante para relatórios mais fáceis de ler.

### `_worksheet_calculation_labels`

Monta uma lista de rótulos de cálculos observados em uma planilha. Essa rotina auxilia a rastrear uso de campos calculados.

### `_custom_sql_by_connection`

Agrupa SQL customizado por conexão de um datasource. O objetivo é preparar uma visão organizada para os relatórios narrativos.

### `_dedupe_external_custom_sql_relations`

Remove duplicidades em blocos de SQL vindos de `.tdsx` externos. Isso evita repetição excessiva nos documentos finais.

### `_dedupe_relationship_maps`

Deduplica relações entre tabelas considerando expressão, origem e campos associados. A rotina evita que o mesmo relacionamento apareça várias vezes.

### `_reference_tokens`

Gera variações textuais possíveis para identificar um objeto. Essa função sustenta busca por referências em diferentes contextos XML.

### `_object_identifier`

Cria um identificador estável para objetos documentados, usando campos disponíveis e um prefixo de fallback. É útil para deduplicação e comparação.

### `_collect_unused_objects`

Calcula o conjunto de objetos sem uso efetivo, incluindo cálculos, parâmetros e datasources. Essa informação alimenta uma seção específica do relatório.

### `_infer_datasource_type`

Classifica a fonte como lógica ou física, com base em sua estrutura e relações observadas. A rotina simplifica a interpretação do modelo de dados.

### `_extract_worksheets`

Extrai planilhas, dependências de datasource, filtros, encodings, estilos e demais atributos relevantes do XML de worksheets.

### `_extract_dashboards`

Extrai dashboards, suas worksheets, zonas, layouts, filtros visíveis e informações visuais. É uma etapa central para documentação analítica do workbook.

### `_extract_stories`

Lê histórias do Tableau quando presentes. O resultado entra no resumo estrutural do workbook.

### `_extract_windows`

Extrai informações de janelas e cartões da interface do workbook. Serve para complementar a visão estrutural e de UI interna.

### `_extract_thumbnails`

Extrai miniaturas embutidas, grava os arquivos correspondentes e registra seus metadados. A rotina contribui para documentação visual do workbook.

### `_extract_visual_tokens`

Percorre o XML buscando tokens visuais de um tipo específico, como cores ou fontes. O resultado é consolidado e deduplicado em etapas posteriores.

### `_parse_filter`

Transforma a estrutura XML de um filtro em um dicionário legível. A rotina identifica campo, classe, visibilidade e outros atributos úteis.

### `_is_context_filter`

Determina se um filtro é de contexto. Essa classificação depende do que o XML do Tableau expõe explicitamente.

### `_humanize_field_reference`

Traduz referências internas de campo para nomes mais amigáveis. O método é aplicado em filtros, dependências e exibições narrativas.

### `_parse_encodings`

Extrai informações de encoding de marks e canais visuais. O resultado ajuda a documentar como a planilha organiza os campos visualmente.

### `_parse_style_rules`

Transforma regras de estilo XML em dicionários estruturados. Essa rotina é reutilizada em estilos globais e locais.

### `_extract_formatted_text`

Extrai conteúdo textual formatado de blocos ricos do XML. É uma rotina auxiliar para preservar texto com estrutura.

### `_collect_colors`

Percorre um elemento XML e coleta valores que parecem cores. Serve como base para a seção de tokens visuais.

### `_collect_fonts`

Percorre um elemento XML e coleta nomes de fontes observados. A saída é usada em dashboards, worksheets e resumos visuais.

### `_display_fonts`

Normaliza a lista de fontes para exibição final. Também introduz o rótulo padrão quando o XML não explicita a tipografia usada.

### `_format_aliases`

Formata aliases em sentenças curtas e prontas para relatório. A ideia é transformar chave e valor em uma representação textual clara.

### `_display_object_name`

Decide o melhor nome para exibir um objeto, priorizando caption e nomes mais humanos. É um utilitário recorrente em seções narrativas.

### `_display_real_name`

Retorna o nome real de um objeto, privilegiando a forma mais útil para leitura técnica.

### `_is_measure_names_column`

Identifica a coluna especial `Measure Names`. Isso permite aplicar tratamento adequado a esse caso específico do Tableau.

### `_display_column_label`

Produz o rótulo final de uma coluna para exibição em relatório. A rotina combina caption, nome real e contexto disponível.

### `_group_datasource_columns`

Agrupa colunas de datasource por tabela ou conjunto lógico. Essa organização melhora a leitura de campos em relatórios extensos.

### `_collect_columns_from_dependencies`

Extrai colunas referenciadas em blocos de dependências. A rotina apoia o entendimento de uso de campos por worksheet.

### `_parse_zones`

Interpreta zonas de layout em dashboards. O resultado ajuda a documentar estrutura visual e composição do painel.

### `_best_effort_xpath`

Tenta produzir um XPath útil para um elemento XML mesmo quando o contexto é limitado. É uma rotina de apoio para o mapa XPath/JSON.

### `generate_xpath_json_map`

Gera os arquivos de mapa XPath/JSON em formato Markdown e JSON. Essa rotina documenta correspondência entre regiões do XML e blocos semânticos da estrutura final.

### `_count_xpath_matches`

Conta quantas ocorrências um XPath encontra no XML carregado. Essa informação é exibida no mapa para indicar densidade ou presença de cada seção.

### `write_outputs`

Coordena a gravação de todos os artefatos finais, respeitando o formato solicitado. Também gera o manifesto e aciona a limpeza de temporários ao final.

### `_cleanup_temporary_outputs`

Remove diretórios e arquivos usados apenas durante o processamento. A rotina preserva os artefatos finais e elimina resíduos intermediários.

## Rotinas de geração de saída

### `_write_json`

Serializa a estrutura consolidada de metadados em JSON indentado. É a representação mais completa para integração e inspeção programática.

### `_write_markdown`

Monta o relatório em Markdown com resumo, workbook, fontes, dashboards, preferências, parâmetros, cálculos e objetos não usados.

### `_write_rtf`

Gera o relatório RTF a partir de blocos documentais abstratos. Essa abordagem evita duplicação de lógica em relação à construção do conteúdo.

### `_write_docx`

Gera o relatório Word. A rotina prepara o documento, insere título, data, índice, quebra de página, rodapé e converte os blocos documentais em parágrafos reais do DOCX.

### `_build_rtf_document`

Converte a lista de blocos abstratos em uma única string RTF. A rotina atua como camada de composição para o formato RTF.

### `_build_document_blocks`

Produz a estrutura canônica do relatório em blocos neutros. Essa é a rotina que concentra o conteúdo documental compartilhado por RTF e DOCX.

### `_doc_paragraph`

Cria um bloco abstrato do tipo parágrafo. É usado como unidade básica do relatório estruturado.

### `_doc_bullet`

Cria um bloco abstrato do tipo bullet. É a base para listas multinível no relatório documental.

### `_doc_code_block`

Cria um bloco abstrato de código ou texto monoespaçado. É usado para fórmulas, diagramas e SQL.

### `_doc_list_block`

Expande uma lista de valores em uma sequência de bullets com rótulo principal e itens subordinados. Essa função reduz repetição ao montar o documento.

### `_render_rtf_block`

Converte um bloco abstrato específico em sintaxe RTF. Atua como ponte entre a estrutura de blocos e o documento textual final.

### `_configure_docx_document`

Aplica configurações globais do Word, como fonte base, espaçamento padrão, propriedades do documento e rodapé.

### `_append_docx_block`

Recebe um bloco abstrato e o transforma em um parágrafo Word real. A rotina decide tratamento para bullets, código e títulos.

### `_apply_docx_run_style`

Define estilo de um `run` do Word, incluindo peso visual, fonte e tamanho. É a base para consistência tipográfica do `.docx`.

### `_set_docx_font_family`

Força a família tipográfica em todos os slots relevantes do XML do Word. Essa rotina existe para evitar que o Word preserve fontes implícitas como Calibri.

### `_apply_docx_paragraph_format`

Aplica o padrão de espaçamento de parágrafo do documento Word. Hoje ele centraliza o comportamento de espaçamento simples, sem espaço antes e com espaço depois.

### `_configure_docx_footer`

Monta o rodapé do `.docx` com nome da pasta de trabalho e paginação dinâmica. A rotina usa alinhamento à esquerda e à direita no mesmo parágrafo.

### `_append_docx_field`

Insere campos dinâmicos simples do Word, como `PAGE` e `NUMPAGES`. É usada no rodapé.

### `_append_docx_toc`

Insere o índice automático do Word logo após a capa inicial do relatório. O índice é configurado para mostrar três níveis principais.

### `_append_docx_complex_field`

Insere campos complexos do Word, como o campo `TOC`. É uma rotina genérica para campos que exigem instruções completas.

### `_docx_heading_style_for_block`

Mapeia estilos lógicos do relatório para estilos de título reconhecidos pelo Word, permitindo que o índice automático funcione corretamente.

### `_rtf_paragraph`

Gera um parágrafo RTF com base em estilo lógico, nível e família tipográfica. Suporta títulos, seções, subtítulos e corpo normal.

### `_rtf_bullet`

Gera um item de lista em RTF com indentação e marcador. É o equivalente RTF dos bullets abstratos do documento.

### `_rtf_code_block`

Gera um bloco monoespaçado em RTF. É usado para fórmulas, SQL e diagramas técnicos.

### `_rtf_list_block`

Expande listas para uma sequência de bullets em RTF. É a contraparte RTF da rotina abstrata `_doc_list_block`.

### `_rtf_escape`

Escapa caracteres especiais e Unicode para sintaxe RTF válida. Sem essa rotina, o documento poderia quebrar ou exibir caracteres incorretos.

### `_looks_like_hex_color`

Valida se um texto se parece com cor hexadecimal. É usada em cenários de apresentação e filtragem visual.

### `_build_connection_diagram`

Monta um diagrama textual simples de conexões, relações e relacionamentos. O objetivo é tornar a estrutura da fonte de dados mais inteligível no relatório.

### `_build_datasources_markdown`

Gera a seção de fontes de dados em Markdown. Inclui campos, conexões, SQL customizado, relacionamentos e metadados publicados.

### `_build_preferences_markdown`

Gera a seção de preferências e paletas em Markdown.

### `_build_parameters_markdown`

Gera a seção de parâmetros em Markdown, incluindo domínio, valor, membros e uso.

### `_build_calculations_markdown`

Gera a seção de campos calculados em Markdown, com foco em origem, uso, impacto e código.

### `_build_unused_objects_markdown`

Gera a seção de objetos não usados em Markdown.

### `_build_dashboards_markdown`

Gera a seção de dashboards em Markdown, incluindo planilhas, filtros, zonas, cores e fontes.

### `_build_visual_tokens_markdown`

Gera a seção de tokens visuais em Markdown, com totais e listas de cores e fontes.

### `_write_excel`

Exporta subconjuntos estruturados dos metadados para múltiplas abas de Excel. Essa saída é útil para exploração tabular e auditoria.

### `_to_frame`

Converte listas de registros em `DataFrame`, com opção de serializar estruturas aninhadas. É a camada de preparo usada pela exportação Excel.

## Rotinas de configuração e CLI

### `load_config`

Carrega e valida o `config/config.json`. A rotina garante que o arquivo exista, contenha JSON válido e tenha estrutura de objeto na raiz.

### `load_path_from_config`

Obtém o caminho padrão do workbook a partir da configuração. Também valida tipo e conteúdo da chave `tableau_path`.

### `load_external_tdsx_paths_from_config`

Lê e normaliza a chave opcional `external_tdsx_paths`. Suporta string única ou lista de caminhos.

### `parse_args`

Define a interface de linha de comando com arquivo opcional e parâmetro `--format`. É a porta de entrada declarativa da execução manual.

### `main`

Executa o fluxo principal do programa. Lê argumentos, resolve configuração, instancia `TableauDoc`, escreve saídas e trata erros conhecidos com mensagem de uso.

## Encadeamento lógico do script

Em alto nível, o script opera em cinco camadas:

1. Entrada e validação: leitura de argumentos, configuração e arquivo-fonte.
2. Carga estrutural: abertura de `.twb` ou extração de `.twbx`.
3. Extração semântica: leitura de datasources, dashboards, parâmetros, cálculos, filtros e tokens visuais.
4. Enriquecimento: humanização de nomes, associação de uso, leitura externa de `.tdsx` e deduplicação.
5. Geração de artefatos: exportação para JSON, Markdown, RTF, DOCX, Excel, mapa XPath/JSON e manifesto.

## Observação final

A organização do código segue uma lógica de transformação progressiva: primeiro o XML bruto é carregado, depois é convertido em estruturas Python, em seguida essas estruturas são enriquecidas e, por fim, transformadas em documentos legíveis. Essa separação torna o script mais previsível para manutenção, mais auditável e mais fácil de estender com novos formatos de saída.
