param([switch]$Template)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$voice = New-Object -ComObject SAPI.SpVoice
$voice.GetVoices() | ForEach-Object { $_.GetDescription() }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    if ($Template) {
        $file = (Get-ChildItem -LiteralPath (Split-Path $root) -Filter '*.docx')[0].FullName
        $pdf = Join-Path $root 'qa/template.pdf'
    } else {
        $file = Join-Path $root 'deliverables/课堂报告.docx'
        $pdf = Join-Path $root 'qa/report.pdf'
    }
    $doc = $word.Documents.Open($file, $false, $true)
    $doc.ExportAsFixedFormat($pdf,17)
    $doc.Close(0)
} finally { $word.Quit() }
if (!$Template) {
    $ppt = New-Object -ComObject PowerPoint.Application
    try {
        $deck = $ppt.Presentations.Open((Join-Path $root 'deliverables/课堂汇报.pptx'),$true,$false,$false)
        $deck.Export((Join-Path $root 'qa/slides'),'PNG',1600,900)
        $deck.Close()
    } finally { $ppt.Quit() }
}
