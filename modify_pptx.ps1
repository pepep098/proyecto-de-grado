Add-Type -AssemblyName System.IO.Compression.FileSystem

$pptxPath = "c:\Users\pipeh\Downloads\ppi\EBNSP-90 A-BLACK.pptx"
$tempDir   = "c:\Users\pipeh\Downloads\ppi\_pptx_temp"

if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
New-Item -ItemType Directory -Path $tempDir | Out-Null

Write-Host "Extrayendo PPTX..."
[System.IO.Compression.ZipFile]::ExtractToDirectory($pptxPath, $tempDir)

function Set-TextColorBlack {
    param($xmlContent)
    $xml = [xml]$xmlContent
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    $rPrNodes = $xml.SelectNodes("//a:rPr", $ns)
    foreach ($rPr in $rPrNodes) {
        $solidFills = $rPr.SelectNodes("a:solidFill", $ns)
        foreach ($sf in $solidFills) {
            $rPr.RemoveChild($sf) | Out-Null
        }
        $solidFill = $xml.CreateElement("a", "solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main")
        $srgbClr = $xml.CreateElement("a", "srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main")
        $srgbClr.SetAttribute("val", "000000")
        $solidFill.AppendChild($srgbClr) | Out-Null
        if ($rPr.HasChildNodes) {
            $rPr.InsertBefore($solidFill, $rPr.FirstChild) | Out-Null
        } else {
            $rPr.AppendChild($solidFill) | Out-Null
        }
    }
    return $xml.OuterXml
}

function Increase-FontSize {
    param($xmlContent, [int]$increasePt = 4)
    $xml = [xml]$xmlContent
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    $increaseVal = $increasePt * 100
    $rPrNodes = $xml.SelectNodes("//a:rPr[@sz]", $ns)
    $count = 0
    foreach ($rPr in $rPrNodes) {
        $currentSz = [int]$rPr.GetAttribute("sz")
        $newSz = $currentSz + $increaseVal
        $rPr.SetAttribute("sz", $newSz.ToString())
        $count++
    }
    Write-Host "    -> $count text elements increased by ${increasePt}pt"
    return $xml.OuterXml
}

$slidesDir = Join-Path $tempDir "ppt\slides"
$slideFiles = Get-ChildItem -Path $slidesDir -Filter "slide*.xml" | Sort-Object {
    [int]($_.BaseName -replace 'slide','')
}

Write-Host ""
Write-Host "Total slides found: $($slideFiles.Count)"

foreach ($slideFile in $slideFiles) {
    $slideNum = [int]($slideFile.BaseName -replace 'slide','')
    Write-Host ""
    Write-Host "Processing slide $slideNum ($($slideFile.Name))..."
    $content = Get-Content -Path $slideFile.FullName -Raw -Encoding UTF8
    if ($slideNum -eq 1) {
        Write-Host "  Setting text color to BLACK..."
        $content = Set-TextColorBlack -xmlContent $content
        Write-Host "  Done - color set to black"
    } elseif ($slideNum -ge 2 -and $slideNum -le 10) {
        Write-Host "  Increasing font size by 4pt..."
        $content = Increase-FontSize -xmlContent $content -increasePt 4
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($slideFile.FullName, $content, $utf8NoBom)
}

Write-Host ""
Write-Host "Repackaging PPTX..."

$tempOutput = $pptxPath + ".new.pptx"
if (Test-Path $tempOutput) { Remove-Item -Force $tempOutput }

[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $tempOutput)

Remove-Item -Force $pptxPath
Rename-Item -Path $tempOutput -NewName (Split-Path $pptxPath -Leaf)

Remove-Item -Recurse -Force $tempDir

Write-Host ""
Write-Host "SUCCESS! Presentation saved: $pptxPath"
Write-Host "  - Slide 1: text in BLACK"
Write-Host "  - Slides 2-10: font size increased by 4pt"
