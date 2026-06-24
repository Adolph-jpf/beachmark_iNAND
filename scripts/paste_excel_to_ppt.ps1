param(
    [string]$RulePath = "Rule_list.xlsx",
    [string]$ReportPath = "",
    [string]$PptPath = "SDSS INAND YIELD WW45_2026_benchmark.pptx",
    [string]$OutputPptPath = "",
    [string]$ReportWeekLabel = "",
    [double]$LeftInch = 0.2,
    [double]$TopInch = 1.5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertTo-NormalizedName([string]$s) {
    if ($null -eq $s) { return "" }
    $x = $s.Trim().ToLowerInvariant()
    $x = [regex]::Replace($x, "\s+", " ")
    return $x
}

function Update-TextRangeWeek($textRange, [string]$reportWeekLabel) {
    if ([string]::IsNullOrWhiteSpace($reportWeekLabel) -or $null -eq $textRange) {
        return 0
    }

    $text = [string]$textRange.Text
    if ([string]::IsNullOrWhiteSpace($text)) {
        return 0
    }

    $matches = [regex]::Matches($text, "W\d{1,2}'\d{2}") |
        ForEach-Object { $_.Value } |
        Sort-Object -Unique

    $changed = 0
    foreach ($old in $matches) {
        if ($old -eq $reportWeekLabel) { continue }
        try {
            $textRange.Replace($old, $reportWeekLabel, 0, 0, 0) | Out-Null
        }
        catch {
            $textRange.Text = [regex]::Replace([string]$textRange.Text, [regex]::Escape($old), $reportWeekLabel)
        }
        $changed++
    }
    return $changed
}

function Update-ShapeWeekText($shape, [string]$reportWeekLabel) {
    $changed = 0

    try {
        if ($shape.Type -eq 6) {
            for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
                $changed += Update-ShapeWeekText $shape.GroupItems.Item($i) $reportWeekLabel
            }
        }
    }
    catch {}

    try {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
            $changed += Update-TextRangeWeek $shape.TextFrame.TextRange $reportWeekLabel
        }
    }
    catch {}

    try {
        if ($shape.HasTable -eq -1) {
            for ($r = 1; $r -le $shape.Table.Rows.Count; $r++) {
                for ($c = 1; $c -le $shape.Table.Columns.Count; $c++) {
                    $cellShape = $shape.Table.Cell($r, $c).Shape
                    if ($cellShape.TextFrame.HasText -eq -1) {
                        $changed += Update-TextRangeWeek $cellShape.TextFrame.TextRange $reportWeekLabel
                    }
                }
            }
        }
    }
    catch {}

    return $changed
}

function Update-ReportWeekText($presentation, [string]$reportWeekLabel) {
    if ([string]::IsNullOrWhiteSpace($reportWeekLabel)) {
        return 0
    }

    $changed = 0
    foreach ($slide in $presentation.Slides) {
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $changed += Update-ShapeWeekText $slide.Shapes.Item($i) $reportWeekLabel
        }
    }
    return $changed
}

function Remove-OldBenchShapes($slide, [double]$leftPt, [double]$topPt) {
    # mso shape type constants
    $msoPicture = 13
    $msoLinkedPicture = 11
    $msoEmbeddedOLEObject = 7
    $msoLinkedOLEObject = 10
    $msoTable = 19

    $deleted = 0
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $sh = $slide.Shapes.Item($i)
        $remove = $false

        # 1) 优先删我们自己贴进去并打过标记的对象
        $nm = [string]$sh.Name
        $alt = [string]$sh.AlternativeText
        if ($alt -eq "AUTO_BENCH_TABLE" -or $nm.StartsWith("AUTO_BENCH_")) {
            $remove = $true
        }
        else {
            # 2) 兼容历史未打标记对象: 目标粘贴区域内的大图/表对象也删掉
            #    避免每次手工删旧表
            $inTargetArea = $false
            $looksLikeTableImage = $false
            $isPasteType = $false
            try {
                $inTargetArea = ($sh.Left -ge ($leftPt - 8)) -and ($sh.Top -ge ($topPt - 8))
                $looksLikeTableImage = ($sh.Width -ge 200) -and ($sh.Height -ge 80)
                $t = $sh.Type
                $isPasteType = @($msoPicture, $msoLinkedPicture, $msoEmbeddedOLEObject, $msoLinkedOLEObject, $msoTable) -contains $t
            }
            catch {
                $inTargetArea = $false
                $looksLikeTableImage = $false
                $isPasteType = $false
            }
            if ($inTargetArea -and $looksLikeTableImage -and $isPasteType) {
                $remove = $true
            }
        }

        if ($remove) {
            $sh.Delete()
            $deleted++
        }
    }
    return $deleted
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $latest = Get-ChildItem -Path "output" -Filter "INAND_weekly_benchmark*.xlsx" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $latest) {
        throw "No output Excel found: output\\INAND_weekly_benchmark*.xlsx"
    }
    $ReportPath = $latest.FullName
}

$RulePath = (Resolve-Path $RulePath).Path
$ReportPath = (Resolve-Path $ReportPath).Path
$PptPath = (Resolve-Path $PptPath).Path
if ([string]::IsNullOrWhiteSpace($OutputPptPath)) {
    $OutputPptPath = $PptPath
}
else {
    $outParent = Split-Path -Parent $OutputPptPath
    if (-not [string]::IsNullOrWhiteSpace($outParent)) {
        New-Item -ItemType Directory -Force -Path $outParent | Out-Null
    }
    if ((Test-Path $OutputPptPath) -and ((Resolve-Path $OutputPptPath).Path -ne $PptPath)) {
        Remove-Item $OutputPptPath -Force
    }
    if ((-not (Test-Path $OutputPptPath)) -or ((Resolve-Path $OutputPptPath).Path -ne $PptPath)) {
        Copy-Item -Path $PptPath -Destination $OutputPptPath -Force
    }
    $OutputPptPath = (Resolve-Path $OutputPptPath).Path
}

Write-Host "Rule  : $RulePath"
Write-Host "Excel : $ReportPath"
Write-Host "PPT template: $PptPath"
Write-Host "PPT output  : $OutputPptPath"

$excel = $null
$ppt = $null
$ruleWb = $null
$reportWb = $null
$pres = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $ruleWb = $excel.Workbooks.Open($RulePath)
    $orderWs = $ruleWb.Worksheets.Item("Sheet1")
    $lastRow = $orderWs.Cells($orderWs.Rows.Count, 1).End(-4162).Row  # xlUp

    $orders = @()
    for ($r = 2; $r -le $lastRow; $r++) {
        $name = [string]$orderWs.Cells($r, 1).Value2
        $slide = $orderWs.Cells($r, 2).Value2
        if ([string]::IsNullOrWhiteSpace($name) -or $null -eq $slide) { continue }
        $orders += [pscustomobject]@{
            Product = $name
            Slide = [int]$slide
        }
    }
    $ruleWb.Close($false)
    $ruleWb = $null

    $reportWb = $excel.Workbooks.Open($ReportPath)
    $sheetMap = @{}
    foreach ($ws in $reportWb.Worksheets) {
        $sheetMap[(ConvertTo-NormalizedName $ws.Name)] = $ws
    }

    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $pres = $ppt.Presentations.Open($OutputPptPath, $false, $false, $false)
    if (-not [string]::IsNullOrWhiteSpace($ReportWeekLabel)) {
        $changedWeekText = Update-ReportWeekText $pres $ReportWeekLabel
        Write-Host "Updated report week text to $ReportWeekLabel ($changedWeekText shape/text range update(s))"
    }

    $leftPt = $LeftInch * 72.0
    $topPt = $TopInch * 72.0
    $slideW = $pres.PageSetup.SlideWidth
    $slideH = $pres.PageSetup.SlideHeight
    $maxW = $slideW - $leftPt - (0.2 * 72.0)
    $maxH = $slideH - $topPt - (0.2 * 72.0)

    $pasted = 0
    foreach ($o in $orders) {
        if ($o.Slide -lt 1 -or $o.Slide -gt $pres.Slides.Count) {
            Write-Warning "Skip [$($o.Product)]: slide index $($o.Slide) out of range"
            continue
        }

        $key = ConvertTo-NormalizedName $o.Product
        if (-not $sheetMap.ContainsKey($key)) {
            Write-Warning "Skip [$($o.Product)]: sheet not found in report workbook"
            continue
        }

        $ws = $sheetMap[$key]
        $used = $ws.UsedRange
        if ($used.Rows.Count -le 1 -or $used.Columns.Count -lt 1) {
            Write-Host "Skip [$($o.Product)]: sheet has no data"
            continue
        }

        $slide = $pres.Slides.Item($o.Slide)
        $removed = Remove-OldBenchShapes $slide $leftPt $topPt
        if ($removed -gt 0) {
            Write-Host "Removed $removed old table shape(s) on slide $($o.Slide)"
        }

        # Copy as picture from Excel and paste to PPT
        $shape = $null
        $pastedOk = $false
        for ($attempt = 1; $attempt -le 5 -and -not $pastedOk; $attempt++) {
            try {
                $used.CopyPicture(1, -4147) | Out-Null   # xlScreen=1, xlPicture=-4147
                Start-Sleep -Milliseconds (150 * $attempt)
                try {
                    $shapeRange = $slide.Shapes.PasteSpecial(2)  # ppPasteEnhancedMetafile
                }
                catch {
                    $shapeRange = $slide.Shapes.Paste()  # fallback
                }
                $shape = $shapeRange.Item(1)
                $pastedOk = $true
            }
            catch {
                if ($attempt -eq 5) { throw }
                Start-Sleep -Milliseconds 250
            }
        }
        $shape.AlternativeText = "AUTO_BENCH_TABLE"
        $safeName = ($ws.Name -replace "[^A-Za-z0-9_]", "_")
        $shape.Name = "AUTO_BENCH_$safeName"
        $shape.LockAspectRatio = -1
        $shape.Left = $leftPt
        $shape.Top = $topPt

        if ($shape.Width -gt $maxW) {
            $shape.Width = $maxW
        }
        if ($shape.Height -gt $maxH) {
            $shape.Height = $maxH
        }

        $pasted++
        Write-Host "Pasted: [$($o.Product)] -> slide $($o.Slide)"
    }

    $pres.Save()
    Write-Host "Done, pasted $pasted table(s) to: $OutputPptPath"
}
finally {
    if ($null -ne $pres) { $pres.Close() }
    if ($null -ne $ppt) { $ppt.Quit() }

    if ($null -ne $reportWb) { $reportWb.Close($false) }
    if ($null -ne $ruleWb) { $ruleWb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
