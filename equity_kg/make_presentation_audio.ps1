param(
    [Parameter(Mandatory=$true)][string]$ContentPath,
    [string]$QaDirectory = ''
)
$ErrorActionPreference = 'Stop'
if (-not $QaDirectory) { $QaDirectory = Join-Path $PSScriptRoot 'qa/ppt-revision' }
$content = Get-Content -LiteralPath $ContentPath -Raw -Encoding UTF8 | ConvertFrom-Json
$audioDir = Join-Path $QaDirectory 'audio'
[void](New-Item -ItemType Directory -Path $audioDir -Force)
$voice = New-Object -ComObject SAPI.SpVoice
$files = @()
try {
    $chinese = $voice.GetVoices() | Where-Object { $_.GetDescription() -like '*Huihui*' } | Select-Object -First 1
    if ($null -eq $chinese) { throw 'Chinese Huihui voice is unavailable.' }
    $voice.Voice = $chinese
    $voice.Rate = 2
    foreach ($slide in $content.slides) {
        $audio = Join-Path $audioDir ('voice-{0}.wav' -f $slide.slide)
        $stream = New-Object -ComObject SAPI.SpFileStream
        try {
            $stream.Open($audio, 3, $false)
            $voice.AudioOutputStream = $stream
            [void]$voice.Speak($slide.narration)
        } finally {
            $stream.Close()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($stream)
        }
        $files += @{ slide=[int]$slide.slide; sha256=(Get-FileHash -LiteralPath $audio -Algorithm SHA256).Hash.ToLower() }
        Write-Output ('Narration saved for slide ' + $slide.slide)
    }
    @{ content_sha256=(Get-FileHash -LiteralPath $ContentPath -Algorithm SHA256).Hash.ToLower(); audio=$files } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $QaDirectory 'audio-manifest.json') -Encoding UTF8
} finally {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($voice)
}
