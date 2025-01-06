# Scripts do PowerShell
### netstat_abno_to_excel.ps1 - Coloca o output do comando netstat -abno diretamente em uma planilha de excel
- pergunta o nome do arquivo para salvar e o local
- formata sem linhas em branco.

<br>

### verificaCheckSum.ps1 - Verificador hash de arquivos:
- Pergunta o caminho do arquivo
- Pergunta o tipo de hash (MD5, SHA1, SHA256, SHA512)
- Pergunta o código hash conhecido do arquivo.
- Mostra na tela a comparação dos hash

<br>

### backupOutlook.ps1 - backup dos arquivos .pst e .ost:
- Modifique a variável $sourcePath pelo caminho onde está os arquivos .pst e .ost
- Modifique a variável $backupPath pelo caminho onde será salvo os arquivos .pst e .ost
