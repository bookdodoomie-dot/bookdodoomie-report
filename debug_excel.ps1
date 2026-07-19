# debug_excel.ps1
$zipFile = "꿈동산 엑셀 데이터.xlsx"
$tempZip = "temp.zip"
$tempDir = "temp_excel"
$log = @()

$log += "Checking if $zipFile exists..."
if (Test-Path $zipFile) {
    $log += "Yes, $zipFile exists!"
} else {
    $log += "No, $zipFile does not exist!"
}

try {
    if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
    if (Test-Path $tempZip) { Remove-Item -Force $tempZip }
    
    $log += "Copying $zipFile to $tempZip..."
    Copy-Item $zipFile $tempZip -ErrorAction Stop
    $log += "Copy completed!"
    
    $log += "Expanding archive..."
    Expand-Archive $tempZip $tempDir -ErrorAction Stop
    $log += "Archive expanded successfully!"
    
    $log += "Listing expanded files..."
    $files = Get-ChildItem -Path $tempDir -Recurse | Select-Object -ExpandProperty FullName
    foreach ($f in $files) {
        $log += $f
    }
} catch {
    $log += "Error occurred: $_"
}

$log | Out-File "debug_log.txt" -Encoding utf8
