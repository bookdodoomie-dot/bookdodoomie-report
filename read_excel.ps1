# read_excel.ps1
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $scriptDir) { $scriptDir = $PSScriptRoot }
if (-not $scriptDir) { $scriptDir = "C:\Users\goal\.gemini\antigravity\scratch\bukdotum-report" }

$zipFile = Join-Path $scriptDir "꿈동산 엑셀 데이터.xlsx"
$tempZip = Join-Path $scriptDir "temp.zip"
$tempDir = Join-Path $scriptDir "temp_excel"

Write-Output "Script directory: $scriptDir"
Write-Output "Target file: $zipFile"
Write-Output "Target exists: $(Test-Path $zipFile)"

if (Test-Path $zipFile) {
    if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
    if (Test-Path $tempZip) { Remove-Item -Force $tempZip }

    Copy-Item $zipFile $tempZip
    Expand-Archive $tempZip $tempDir

    $sharedStrings = @()
    $sharedStringsFile = Join-Path $tempDir "xl/sharedStrings.xml"
    if (Test-Path $sharedStringsFile) {
        [xml]$xml = Get-Content $sharedStringsFile -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        $tNodes = $xml.SelectNodes("//x:t", $ns)
        foreach ($node in $tNodes) {
            $sharedStrings += $node.InnerText
        }
    }

    $sheetFile = Join-Path $tempDir "xl/worksheets/sheet1.xml"
    $output = @()
    if (Test-Path $sheetFile) {
        [xml]$xml = Get-Content $sheetFile
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        
        $rows = $xml.SelectNodes("//x:row", $ns)
        foreach ($row in $rows) {
            $rowData = @()
            $cells = $row.SelectNodes("x:c", $ns)
            foreach ($cell in $cells) {
                $valNode = $cell.SelectSingleNode("x:v", $ns)
                $val = ""
                if ($valNode) {
                    $valIndex = $valNode.InnerText
                    if ($cell.t -eq "s") {
                        $val = $sharedStrings[[int]$valIndex]
                    } else {
                        $val = $valIndex
                    }
                }
                $rowData += $val
            }
            $output += ($rowData -join "`t")
        }
    }

    $output | Write-Output

    # Clean up
    if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
    if (Test-Path $tempZip) { Remove-Item -Force $tempZip }
} else {
    Write-Output "File not found."
}
