param(
    [Parameter(Mandatory=$true)][string]$ContentPath,
    [string]$QaDirectory = ''
)
$ErrorActionPreference = 'Stop'
if (-not $QaDirectory) { $QaDirectory = Join-Path $PSScriptRoot 'qa/ppt-revision' }
$plan = Get-Content -LiteralPath $ContentPath -Raw -Encoding UTF8 | ConvertFrom-Json
$source = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $plan.source))
$target = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $plan.output))
if ($source -eq $target) { throw 'The revised deck must be a separate copy.' }
$originalHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$slidesDir = Join-Path $QaDirectory 'slides'
[void](New-Item -ItemType Directory -Path $slidesDir -Force)
$ppt = New-Object -ComObject PowerPoint.Application
$ownedApplication = ($ppt.Presentations.Count -eq 0)
$deck = $null
$audit = @()
try {
    $deck = $ppt.Presentations.Open($source, -1, 0, 0)
    if ($deck.Slides.Count -ne $plan.slides.Count) { throw 'Slide count differs from the content plan.' }
    foreach ($entry in $plan.slides) {
        $slide = $deck.Slides.Item([int]$entry.slide)
        foreach ($replacement in $entry.replacements) {
            $old = [string]$replacement[0]
            $new = [string]$replacement[1]
            $matches = @($slide.Shapes | Where-Object {
                $_.HasTextFrame -eq -1 -and $_.TextFrame.HasText -eq -1 -and
                ($_.TextFrame.TextRange.Text -replace "`r`n?", "`n") -eq $old
            })
            if ($matches.Count -ne 1) { throw ('Expected one exact text match on slide '+$entry.slide+': '+$old) }
            $shape = $matches[0]
            $beforeFont = [double]$shape.TextFrame.TextRange.Font.Size
            $beforeName = [string]$shape.TextFrame.TextRange.Font.Name
            $beforeBold = [int]$shape.TextFrame.TextRange.Font.Bold
            $beforeItalic = [int]$shape.TextFrame.TextRange.Font.Italic
            $range = $shape.TextFrame.TextRange
            [void]$range.Replace(($old -replace "`n", "`r"), ($new -replace "`n", "`r"), 0, -1, 0)
            # Replace can introduce the Office theme's CJK font despite retaining
            # Font.Name. Keep the resolved source family for both script systems.
            $shape.TextFrame.TextRange.Font.NameFarEast = $beforeName
            $shape.TextFrame.TextRange.Font.Bold = $beforeBold
            $shape.TextFrame.TextRange.Font.Italic = $beforeItalic
            if (($shape.TextFrame.TextRange.Text -replace "`r`n?", "`n") -ne $new) { throw 'Text replacement failed.' }
            if ([double]$shape.TextFrame.TextRange.Font.Size -ne $beforeFont -or [string]$shape.TextFrame.TextRange.Font.Name -ne $beforeName) {
                throw ('Unexpected font change on slide '+$entry.slide)
            }
            $audit += [ordered]@{ slide=[int]$entry.slide; shape_id=$shape.Id; font=$beforeName; font_size=$beforeFont; text=$new }
        }
        $body = @($slide.NotesPage.Shapes | Where-Object { $_.Type -eq 14 -and $_.PlaceholderFormat.Type -eq 2 })
        if ($body.Count -ne 1) { throw ('Missing notes body on slide '+$entry.slide) }
        $body[0].TextFrame.TextRange.Text = $entry.narration+"`r`r[Sources]`r"+($entry.sources -join "`r")+"`r[/Sources]"
    }
    $deck.SaveAs($target, 24)
    $deck.Export($slidesDir, 'PNG', 1600, 900)
    $bounds = @()
    foreach ($slide in $deck.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame.HasText -ne -1) { continue }
            $range = $shape.TextFrame.TextRange
            $bounds += [ordered]@{
                slide=$slide.SlideIndex; shape_id=$shape.Id; text=$range.Text
                left=[double]$shape.Left; top=[double]$shape.Top; width=[double]$shape.Width; height=[double]$shape.Height
                bound_left=[double]$range.BoundLeft; bound_top=[double]$range.BoundTop
                bound_width=[double]$range.BoundWidth; bound_height=[double]$range.BoundHeight
                font_size=[double]$range.Font.Size; font=[string]$range.Font.Name
                word_wrap=[int]$shape.TextFrame.WordWrap; auto_size=[int]$shape.TextFrame.AutoSize
            }
        }
    }
    [ordered]@{ source_sha256=$originalHash; target_sha256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash; content_sha256=(Get-FileHash -LiteralPath $ContentPath -Algorithm SHA256).Hash; edits=$audit; text_bounds=$bounds } |
        ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $QaDirectory 'powerpoint-audit.json') -Encoding UTF8
    $talk = '# 课堂汇报讲稿（修订版）'+"`n`n"+(($plan.slides | ForEach-Object { '## 第'+$_.slide+'页 '+$_.title+"`n`n"+$_.narration }) -join "`n`n")
    Set-Content -LiteralPath (Join-Path $PSScriptRoot 'deliverables/汇报讲稿_修订版.md') -Value $talk -Encoding UTF8
    Write-Output ('Saved revised deck and rendered '+$deck.Slides.Count+' slides: '+$target)
} finally {
    if ($null -ne $deck) { $deck.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck) }
    if ($ownedApplication -and $ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt)
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $originalHash) { throw 'Source deck was unexpectedly modified.' }
}
