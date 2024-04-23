# Verifica se o script está sendo executado como administrador
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Este script precisa ser executado como administrador."
    Write-Host "Por favor, execute o PowerShell como administrador e tente novamente."
    Exit
}

# Executa o comando netstat -abno e captura a saída
$netstatOutput = netstat -abno

# Divide a saída em linhas
$netstatLines = $netstatOutput -split '\r\n'

# Cria um novo objeto Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Define o cabeçalho da tabela do Excel
$worksheet.Cells.Item(1,1) = "Protocolo"
$worksheet.Cells.Item(1,2) = "Endereço Local"
$worksheet.Cells.Item(1,3) = "Endereço Remoto"
$worksheet.Cells.Item(1,4) = "Estado"
$worksheet.Cells.Item(1,5) = "PID"
$worksheet.Cells.Item(1,6) = "Processo"

# Preenche a tabela do Excel com os dados do netstat
$row = 2
for ($i=0; $i -lt $netstatLines.Count; $i++) {
    $line = $netstatLines[$i]
    # Verifica se a linha contém informações relevantes
    if ($line -match '^  (TCP|UDP)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d+)$') {
        $proto = $matches[1]
        $localAddress = $matches[2]
        $foreignAddress = $matches[3]
        $state = $matches[4]
        $pids = $matches[5]

        # Busca o processo na próxima linha
        $nextLine = $netstatLines[$i+1]
        if ($nextLine -match '^\s+(.+)$') {
            $process = $matches[1]
        } else {
            $process = "N/A"
        }

        $worksheet.Cells.Item($row,1) = $proto
        $worksheet.Cells.Item($row,2) = $localAddress
        $worksheet.Cells.Item($row,3) = $foreignAddress
        $worksheet.Cells.Item($row,4) = $state
        $worksheet.Cells.Item($row,5) = $pids
        $worksheet.Cells.Item($row,6) = $process

        $row++
    }
}

# Ajusta a largura das colunas
$worksheet.Columns.AutoFit() 

# Pergunta onde salvar o arquivo e qual o nome
$fileName = Read-Host "Digite o nome do arquivo para salvar (pressione Enter para usar o nome padrão)"
if ($fileName -eq "") {
    $fileName = "netstat_output.xlsx"
}

$saveLocation = Read-Host "Digite o caminho para salvar o arquivo (pressione Enter para salvar no diretório atual)"
if ($saveLocation -eq "") {
    $saveLocation = $PSScriptRoot
}

# Monta o caminho completo para salvar o arquivo
$savePath = Join-Path -Path $saveLocation -ChildPath $fileName

# Salva o arquivo Excel
$workbook.SaveAs($savePath)

# Fecha o Excel
$excel.Quit()

Write-Host "Arquivo salvo em: $savePath"
