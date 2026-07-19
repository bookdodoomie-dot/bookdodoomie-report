# inspect_comments.ps1
$xlsxPath = "C:\Users\goal\.gemini\antigravity\scratch\bukdotum-report\꿈동산 엑셀 데이터.xlsx"
$tempDir = "C:\Users\goal\.gemini\antigravity\scratch\bukdotum-report\temp_inspect"
$outputFile = "C:\Users\goal\.gemini\antigravity\scratch\bukdotum-report\excel_inspection.txt"

$log = @()

try {
    if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
    
    $log += "Expanding XLSX archive..."
    Expand-Archive $xlsxPath $tempDir -Force
    
    # 1. Parse shared strings
    $sharedStrings = @()
    $ssFile = Join-Path $tempDir "xl/sharedStrings.xml"
    if (Test-Path $ssFile) {
        [xml]$xml = Get-Content $ssFile -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        $tNodes = $xml.SelectNodes("//x:t", $ns)
        foreach ($node in $tNodes) {
            $sharedStrings += $node.InnerText
        }
    }
    $log += "Loaded $($sharedStrings.Count) shared strings."

    # 2. Get sheets mapping
    $wbFile = Join-Path $tempDir "xl/workbook.xml"
    $sheets = @()
    if (Test-Path $wbFile) {
        [xml]$xml = Get-Content $wbFile
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        $sheetNodes = $xml.SelectNodes("//x:sheet", $ns)
        foreach ($node in $sheetNodes) {
            $sheets += [PSCustomObject]@{
                Name = $node.name
                Id = $node.GetAttribute("r:id")
            }
        }
    }

    $wbRelsFile = Join-Path $tempDir "xl/_rels/workbook.xml.rels"
    $rels = @{}
    if (Test-Path $wbRelsFile) {
        [xml]$xml = Get-Content $wbRelsFile
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        $relNodes = $xml.SelectNodes("//x:Relationship", $ns)
        foreach ($node in $relNodes) {
            $rels[$node.Id] = $node.Target
        }
    }

    # 3. Read first sheet XML and print all non-empty text cells
    if ($sheets.Count -gt 0) {
        $firstSheet = $sheets[0]
        $targetPath = $rels[$firstSheet.Id]
        $sheetFile = Join-Path $tempDir "xl/$targetPath"
        
        $log += "========================================"
        $log += "INSPECTING SHEET: $($firstSheet.Name) ($targetPath)"
        $log += "========================================"
        
        if (Test-Path $sheetFile) {
            [xml]$xml = Get-Content $sheetFile
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
            
            $cells = $xml.SelectNodes("//x:c", $ns)
            foreach ($c in $cells) {
                $ref = $c.GetAttribute("r")
                $vNode = $c.SelectSingleNode("x:v", $ns)
                if ($vNode) {
                    $vText = $vNode.InnerText
                    $val = ""
                    if ($c.GetAttribute("t") -eq "s") {
                        $val = $sharedStrings[[int]$vText]
                    } else {
                        $val = $vText
                    }
                    
                    if ($val -and $val.Trim().Length -gt 0) {
                        # Print cells with text, especially long ones which are likely comments
                        $log += "  Cell $ref : $val"
                    }
                }
            }
        }
    }
} catch {
    $log += "Error occurred: $_"
    $log += $_.ScriptStackTrace
} finally {
    if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue }
    $log | Out-File $outputFile -Encoding utf8
}
