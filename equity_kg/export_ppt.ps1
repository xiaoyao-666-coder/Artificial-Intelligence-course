$ErrorActionPreference='Stop'
$ppt=New-Object -ComObject PowerPoint.Application
try {
    $deck=$ppt.Presentations.Open((Join-Path $PSScriptRoot 'deliverables/课堂汇报.pptx'),$true,$false,$false)
    $deck.Export((Join-Path $PSScriptRoot 'qa/slides'),'PNG',1600,900)
    $deck.Close()
} finally { $ppt.Quit() }
