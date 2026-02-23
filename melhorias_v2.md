# Melhorias Tecnicas - Inventory_v2.ps1 (v2.1.0)

## Contexto da refatoracao
A refatoracao do `Inventory_v2.ps1` foi direcionada para robustez de dados, consistencia estrutural dos CSVs e maior observabilidade da execucao.
O objetivo foi reduzir variabilidade operacional sem alterar o objetivo funcional macro do inventario: coletar dados, gerar artefatos, compactar, enviar por FTP e limpar localmente em sucesso.

## Comparativo v2.0.3 -> v2.1.0
| Eixo | Antes (v2.0.3) | Depois (v2.1.0) | Impacto tecnico |
|---|---|---|---|
| Normalizacao/sanitizacao | Tratamento pontual de formato | Camada dedicada com `Normalize-Text`, `Normalize-Date`, `Normalize-User` | Menor ruido de caracteres, datas e identidade de usuario |
| Escrita de CSV | Escrita linha a linha com `Write-CsvLine` | Buffer por schema com `New-CsvBuffer` + `Add-CsvRowSafe` + `Flush-CsvBuffer` | Escrita previsivel, acoplada a estrutura declarada |
| Validacao de linhas/colunas | Sem validacao estrutural formal | `Validate-CsvRowCount` por arquivo antes de gravar | Reduz divergencia de colunas e quebra de consumo |
| Validacao do payload consolidado | Resultado usado direto apos coleta | `Validate-InventoryResult` para normalizar, filtrar e ordenar colecoes | Integridade de payload antes da serializacao tabular |
| Metricas por etapa | Logs informativos sem metricas de desempenho por etapa | `Invoke-InventoryStep` registra `status`, `duracao_ms`, `item_count` | Maior rastreabilidade e diagnstico operacional |
| Deduplicacao de software/usuarios | Dedupe mais simples e disperso | Dedupe estruturado com chave composta e mapas normalizados | Menor duplicidade e maior consistencia de inventario |

## Funcoes novas e papel tecnico
- `Get-Utf8BomEncoding`: define codificacao UTF-8 com BOM para compatibilidade de leitura em ferramentas que exigem BOM.
- `Write-AllLinesUtf8`: grava arquivos em lote, evitando append incremental fragil e padronizando encoding.
- `Normalize-Text`: sanitiza texto (quebra de linha, separador `;`, aspas), mitigando corrupcao de CSV.
- `Normalize-Date`: converte datas para formato operacional consistente (`dd/MM/yyyy` e `dd/MM/yyyy HH:mm:ss`).
- `Normalize-User`: normaliza identificacao de usuario para comparacao confiavel entre sessoes/perfis.
- `New-CsvSchemaMap`: centraliza contratos de colunas por tipo de CSV.
- `New-CsvBuffer`: cria buffer com schema fixo (path, colunas, header e linhas).
- `Add-CsvRowSafe`: adiciona linhas via mapa de colunas, com controle de cardinalidade de campos.
- `Validate-CsvRowCount`: valida estrutura de cada linha contra schema esperado antes da persistencia.
- `Flush-CsvBuffer`: aplica validacao e grava arquivo final em UTF-8 BOM.
- `Get-StepItemCount`: calcula cardinalidade de retorno por etapa para telemetria de coleta.
- `Validate-InventoryResult`: normaliza resultado consolidado (ordenacao, filtro de nulos e consistencia minima).

## Fluxo de dados refatorado
Fluxo tecnico aplicado na v2.1.0:
coleta -> fallback controlado -> validacao de payload -> montagem de buffers CSV por schema -> validacao estrutural de linhas/colunas -> gravacao UTF-8 BOM -> consolidacao -> upload FTP -> limpeza local.
Esse encadeamento reduz risco de inconsistencias silenciosas e torna a saida mais deterministica, mantendo o mesmo resultado funcional esperado da linha v2.

## Beneficios tecnicos observaveis
- Maior previsibilidade de output por schema declarado.
- Menor chance de quebra por caracteres especiais, delimitadores e formatos heterogeneos.
- Rastreabilidade por etapa com metricas (`duracao_ms`, `status`, `item_count`).
- Melhoria da integridade dos CSVs para consumo downstream (ETL, auditoria e suporte).

## Compatibilidade e comportamento mantido
- Caminho base operacional mantido: `C:\TI`.
- Geracao de ZIP mantida.
- Upload FTP mantido.
- Limpeza local em sucesso mantida.
- Preservacao de artefatos locais em falha mantida.

## Criterios tecnicos de aceite
- CSVs gerados conforme schema declarado.
- Ausencia de divergencia de colunas por linha.
- Logs contendo metricas de etapa.
- Fluxo fim-a-fim preservado (coleta -> CSV -> ZIP -> FTP -> limpeza em sucesso).
