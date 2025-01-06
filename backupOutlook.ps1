# Defina o caminho do diretório de origem e o destino do backup
$sourcePath = "C:\Users\SEUUSUARIO\AppData\Local\Microsoft\Outlook"
$backupPath = "D:\EmailBackup"

Write-Host "Iniciando o processo de backup..."

# Crie o diretório de backup, se não existir
if (!(Test-Path -Path $backupPath)) {
    New-Item -ItemType Directory -Path $backupPath
    Write-Host "Diretório de backup criado em $backupPath"
} else {
    Write-Host "Diretório de backup já existe em $backupPath"
}

# Copie os arquivos .pst e .ost para o diretório de backup
$extensions = @("*.pst", "*.ost")
foreach ($ext in $extensions) {
    $files = Get-ChildItem -Path $sourcePath -Filter $ext
    $totalFiles = $files.Count
    $counter = 0

    foreach ($file in $files) {
        Copy-Item -Path $file.FullName -Destination $backupPath
        $counter++
        Write-Progress -Activity "Copiando arquivos $ext" -Status "$counter de $totalFiles copiados" -PercentComplete (($counter / $totalFiles) * 100)
    }
}

Write-Host "Processo de backup concluído."
