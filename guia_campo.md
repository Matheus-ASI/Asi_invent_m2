# Guia de Campo - Inventario ASI

## Objetivo
Este guia orienta a execucao operacional dos scripts de inventario ASI em campo.
O foco e padronizar coleta, validacao e evidencia de execucao para registro em OS.
Use este documento como procedimento pratico para execucao assistida e tratativa inicial de falhas.

## Scripts suportados
| Script | Formato | Uso recomendado | Versao |
|---|---|---|---|
| `Inventory_v2.ps1` | CSV + ZIP + FTP | Padrao atual para inventario de estacoes | `2.1.0` |
| `Inventory.ps1` | CSV + ZIP + FTP | Legado imediato, manter para compatibilidade operacional | `2.0.3` |

## Pre-requisitos da maquina
- PowerShell 5.1 disponivel e compativel (validar com `$PSVersionTable.PSVersion`).
- Rodar como administrador da máquina 
- Permissao para criar e gravar em `C:\TI`.
- A máquina deve ter conexão com a internet para o enviar o relatório.
- Execucao com politica adequada (`Bypass` no comando remoto quando aplicavel).

## Passo a passo de execucao (padrao recomendado)
1. Abra PowerShell como usuario com administrador da máquina`.
2. Execute o comando remoto recomendado para `Inventory_v2.ps1`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $u='https://raw.githubusercontent.com/Matheus-ASI/Asi_invent_m2/main/Inventory_v2.ps1'; $p=Join-Path $env:TEMP 'Inventory_v2.ps1'; iwr -UseBasicParsing $u -OutFile $p; Unblock-File $p -ErrorAction SilentlyContinue; & $p"
```

3. Alternativa local (arquivo ja baixado):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Inventory_v2.ps1
```

4. Teste rapido de coleta de software (sem fluxo completo):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Inventory_v2.ps1 -TestSoftware
```

5. Aguarde a finalizacao e valide os itens do checklist abaixo.

## Checklist de validacao (apos execucao)
- [ ] Log iniciou e finalizou sem erro fatal.
- [ ] CSVs de categorias foram gerados.
- [ ] ZIP da maquina foi criado.
- [ ] Upload FTP foi concluido.
- [ ] Limpeza local pos-sucesso foi executada.
- [ ] Em caso de falha, os artefatos locais foram preservados para analise.

## Saidas esperadas
- `C:\TI\<NOME_DA_MAQUINA>\*.csv`.
- `C:\TI\<NOME_DA_MAQUINA>.zip`.
- `C:\TI\<NOME_DA_MAQUINA>\logs\inventario.log`.

## Tratativa rapida de falhas comuns
1. Erro de permissao em `C:\TI`
Sintoma: falha ao criar pasta/arquivo no inicio da execucao.
Acao imediata: executar com conta autorizada e validar permissao NTFS/local.

2. Bloqueio por antivirus/EDR
Sintoma: script interrompido, arquivo removido/quarentenado ou bloqueio do `powershell.exe`.
Acao imediata: validar os logs do antivirus/EDR, aplicar excecao temporaria conforme politica interna e reexecutar.

3. Politica de permissao do PowerShell
Sintoma: erro de execucao com mensagem do tipo `running scripts is disabled on this system`.
Acao imediata: executar com `-ExecutionPolicy Bypass` e validar politicas ativas com `Get-ExecutionPolicy -List`.

4. Incompatibilidade de versao do PowerShell
Sintoma: comandos/funcoes com comportamento inconsistente ou erro de execucao.
Acao imediata: confirmar versao com `$PSVersionTable.PSVersion` e executar no PowerShell 5.1.

5. Falha de upload FTP
Sintoma: log para no envio ou retorna erro de conexao/autenticacao.
Acao imediata: validar conexão de internet da máquina e tentar nova execução.

6. Falha de coleta parcial (etapas com WARN)
Sintoma: uma ou mais etapas registram `WARN`, mas script segue.
Acao imediata: registrar etapa afetada na OS e anexar log para analise interna.

7. Script interrompido
Sintoma: janela fecha, timeout externo ou encerramento manual.
Acao imediata: reexecutar o script e confirmar integridade dos artefatos gerados.

8. Endpoint inacessivel
Sintoma: timeout de rede no download RAW ou no upload FTP.
Acao imediata: validar DNS/proxy/firewall local e executar novamente apos normalizacao.

## Registro na OS
Use o bloco abaixo para registrar a execucao:

```text
Script executado:
Horario (inicio/fim):
Resultado (Sucesso/Falha parcial/Falha):
Evidencias:
- Log: C:\TI\<NOME_DA_MAQUINA>\logs\inventario.log
- Arquivos gerados:(nome do zip gerado)
- Observacoes tecnicas:
```
