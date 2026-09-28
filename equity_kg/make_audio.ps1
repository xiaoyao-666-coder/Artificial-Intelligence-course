$ErrorActionPreference = 'Stop'
$texts = Get-Content -Raw -Encoding utf8 (Join-Path $PSScriptRoot 'qa/narration.json') | ConvertFrom-Json
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Voice = $voice.GetVoices() | Where-Object { $_.GetDescription() -like '*Huihui*' } | Select-Object -First 1
$voice.Rate = 2
for ($i=0; $i -lt $texts.Count; $i++) {
    $stream = New-Object -ComObject SAPI.SpFileStream
    $stream.Open((Join-Path $PSScriptRoot ('qa/voice-{0}.wav' -f ($i+1))),3,$false)
    $voice.AudioOutputStream = $stream
    [void]$voice.Speak($texts[$i])
    $stream.Close()
}
