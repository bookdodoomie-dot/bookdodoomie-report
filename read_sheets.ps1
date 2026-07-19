# read_sheets.ps1
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = "C:\Users\goal\.gemini\antigravity\scratch\bukdotum-report"
$xlsxPath = Join-Path $scriptDir "꿈동산 엑셀 데이터.xlsx"
$tempExtractDir = Join-Path $scriptDir "temp_extract"
$outputFile = Join-Path $scriptDir "sheets_info.txt"

$log = @()

try {
    $log += "Checking Excel path: $xlsxPath"
    if (-not (Test-Path $xlsxPath)) {
        throw "Excel file does not exist at $xlsxPath"
    }

    if (Test-Path $tempExtractDir) {
        $log += "Cleaning existing temp directory..."
        Remove-Item -Recurse -Force $tempExtractDir -ErrorAction SilentlyContinue
    }

    $log += "Extracting $xlsxPath..."
    [System.IO.Compression.ZipFile]::ExtractToDirectory($xlsxPath, $tempExtractDir)

    # 1. Parse sharedStrings.xml
    $sharedStrings = @()
    $ssFile = Join-Path $tempExtractDir "xl/sharedStrings.xml"
    if (Test-Path $ssFile) {
        $log += "Parsing sharedStrings.xml..."
        [xml]$xml = Get-Content $ssFile -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        $tNodes = $xml.SelectNodes("//x:t", $ns)
        foreach ($node in $tNodes) {
            $sharedStrings += $node.InnerText
        }
    }
    $log += "Shared Strings Count: $($sharedStrings.Count)"

    # 2. Parse workbook.xml to get sheet names
    $wbFile = Join-Path $tempExtractDir "xl/workbook.xml"
    $sheets = @()
    if (Test-Path $wbFile) {
        $log += "Parsing workbook.xml..."
        [xml]$xml = Get-Content $wbFile
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
        $sheetNodes = $xml.SelectNodes("//x:sheet", $ns)
        foreach ($node in $sheetNodes) {
            $sheets += [PSCustomObject]@{
                Name = $node.name
                SheetId = $node.sheetId
                Id = $node.GetAttribute("r:id")
            }
        }
    }

    $log += "Sheets found:"
    foreach ($s in $sheets) {
        $log += "  Sheet Name: $($s.Name), SheetId: $($s.SheetId), Id: $($s.Id)"
    }

    # 3. Read specific cells in worksheets
    $wbRelsFile = Join-Path $tempExtractDir "xl/_rels/workbook.xml.rels"
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

    foreach ($sheet in $sheets) {
        $targetPath = $rels[$sheet.Id]
        $sheetFile = Join-Path $tempExtractDir "xl/$targetPath"
        if (Test-Path $sheetFile) {
            $log += "----------------------------------------"
            $log += "Reading Sheet Name: $($sheet.Name) ($targetPath)"
            [xml]$xml = Get-Content $sheetFile
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("x", $xml.DocumentElement.NamespaceURI)
            
            # Check cells D18, F18, H18, J18
            $targetCells = @("D18", "F18", "H18", "J18")
            foreach ($cRef in $targetCells) {
                $cNode = $xml.SelectSingleNode("//x:c[@r='$cRef']", $ns)
                $val = "Empty"
                if ($cNode) {
                    $vNode = $cNode.SelectSingleNode("x:v", $ns)
                    if ($vNode) {
                        $vText = $vNode.InnerText
                        if ($cNode.GetAttribute("t") -eq "s") {
                            $val = $sharedStrings[[int]$vText]
                        } else {
                            $val = $vText
                        }
                    }
                }
                $log += "  Cell $cRef : $val"
            }
            
            # Print Row 18 cells
            $log += "  All cells in Row 18:"
            $row18Node = $xml.SelectSingleNode("//x:row[@r='18']", $ns)
            if ($row18Node) {
                $cells = $row18Node.SelectNodes("x:c", $ns)
                foreach ($c in $cells) {
                    $rAttr = $c.GetAttribute("r")
                    $vNode = $c.SelectSingleNode("x:v", $ns)
                    $vVal = "Empty"
                    if ($vNode) {
                        $vText = $vNode.InnerText
                        if ($c.GetAttribute("t") -eq "s") {
                            $vVal = $sharedStrings[[int]$vText]
                        } else {
                            $vVal = $vText
                        }
                    }
                    $log += "    $rAttr : $vVal"
                }
            } else {
                $log += "    Row 18 not found"
            }
        }
    }
} catch {
    $log += "ERROR OCCURRED: $_"
    $log += $_.ScriptStackTrace
} finally {
    # Clean up
    if (Test-Path $tempExtractDir) {
        Remove-Item -Recurse -Force $tempExtractDir -ErrorAction SilentlyContinue
    }
    # Save log
    $log | Out-File $outputFile -Encoding utf8
}
