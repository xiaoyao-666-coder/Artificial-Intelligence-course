param([string]$DeckPath = '', [string]$QaDirectory = '')
$ErrorActionPreference = 'Stop'
if (-not $DeckPath) { $DeckPath = Join-Path $PSScriptRoot 'deliverables/企业股权知识表示与风险推理_暖色增强版.pptx' }
if (-not $QaDirectory) { $QaDirectory = Join-Path $PSScriptRoot 'qa/ppt-enriched' }
$DeckPath = [IO.Path]::GetFullPath($DeckPath)
$slidesDir = Join-Path $QaDirectory 'slides'
[void](New-Item -ItemType Directory -Path $slidesDir -Force)
$ppt = New-Object -ComObject PowerPoint.Application
$ownedApplication = ($ppt.Presentations.Count -eq 0)
$deck = $null
try {
    $deck = $ppt.Presentations.Open($DeckPath, -1, 0, 0)
    Write-Output ('Opened '+$deck.Slides.Count+' slides for rendering.')
    $bounds = @()
    foreach ($slide in $deck.Slides) {
        $slide.Export((Join-Path $slidesDir ('slide-{0:d2}.png' -f $slide.SlideIndex)), 'PNG', 1600, 900)
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame.HasText -ne -1) { continue }
            $r = $shape.TextFrame.TextRange
            $bounds += [ordered]@{
                slide=$slide.SlideIndex; id=$shape.Id; text=$r.Text
                left=[double]$shape.Left; top=[double]$shape.Top
                width=[double]$shape.Width; height=[double]$shape.Height
                bound_left=[double]$r.BoundLeft; bound_top=[double]$r.BoundTop
                bound_width=[double]$r.BoundWidth; bound_height=[double]$r.BoundHeight
            }
        }
    }
    $pdf = [IO.Path]::ChangeExtension($DeckPath, '.pdf')
    $deck.SaveAs($pdf, 32)
    $sha = [Security.Cryptography.SHA256]::Create()
    $deckHash = [BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($DeckPath))).Replace('-', '')
    $sha.Dispose()
    [ordered]@{
        deck_sha256=$deckHash
        slide_count=$deck.Slides.Count
        text_bounds=$bounds
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $QaDirectory 'render-audit.json') -Encoding UTF8
    Write-Output ('Rendered '+$deck.Slides.Count+' slides and saved '+$pdf)
} finally {
    if ($null -ne $deck) { $deck.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck) }
    if ($ownedApplication -and $ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt)
}
