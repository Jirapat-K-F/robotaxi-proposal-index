# Export every slide as thumbs/<n>.jpg (needs PowerPoint). Usage: powershell -File export_thumbs.ps1 "deck.pptx"
param([string]$deck = "Robotaxi - Proposal Presentation (1).pptx")
$src = (Resolve-Path $deck).Path
$out = Join-Path $PSScriptRoot "thumbs"
New-Item -ItemType Directory -Force $out | Out-Null
Remove-Item "$out\*.jpg" -ErrorAction SilentlyContinue
$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open($src, $true, $false, $false)
foreach ($s in $p.Slides) { $s.Export("$out\$($s.SlideIndex).jpg", "JPG", 960, 540) }
"exported $($p.Slides.Count) slides"
$p.Close(); $pp.Quit()
