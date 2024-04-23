# Solicita o caminho do arquivo
$filePath = Read-Host "Digite ou cole o caminho do arquivo"

# Solicita o tipo de hash
$hashAlgorithm = Read-Host "Digite o tipo de hash (MD5, SHA1, SHA256, SHA512)"

# Verifica se o tipo de hash é válido
if ($hashAlgorithm -notin @("MD5", "SHA1", "SHA256", "SHA512")) {
    Write-Host "Tipo de hash inválido. Por favor, escolha entre MD5, SHA1, SHA256 ou SHA512."
    Exit
}

# Calcula o hash do arquivo com o algoritmo selecionado
$fileHash = Get-FileHash -Algorithm $hashAlgorithm $filePath | Select-Object -ExpandProperty Hash

# Solicita o hash conhecido
$knownChecksum = Read-Host "Digite ou cole o hash conhecido"

# Compara o hash calculado com o hash conhecido
if ($fileHash -eq $knownChecksum) {
    Write-Host "O arquivo é autêntico."
    Write-Host "Hash do arquivo local/save: $fileHash"
    Write-Host "Hash conhecido do arquivo : $knownChecksum"
} else {
    Write-Host "O arquivo foi modificado ou corrompido."
    Write-Host "Hash do arquivo local/save: $fileHash"
    Write-Host "Hash conhecido do arquivo : $knownChecksum"
}

