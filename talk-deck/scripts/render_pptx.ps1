param([string]$Pptx, [string]$OutDir, [string]$Slides = "")
# Export slides to PNG (1600x900) via PowerPoint COM.
# Opens a temp COPY read-only without a window and only Close()s it. Never $app.Quit():
# PowerPoint is a single instance, Quit would also close the user's open decks.
# Usage: powershell -File render_pptx.ps1 -Pptx deck.pptx -OutDir png_dir [-Slides "3,8,12"]
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $OutDir | Out-Null
$tmp = Join-Path $env:TEMP ("render_" + [guid]::NewGuid().ToString() + ".pptx")
Copy-Item $Pptx $tmp
$app = New-Object -ComObject PowerPoint.Application
$p = $app.Presentations.Open($tmp, -1, 0, 0)   # ReadOnly, Untitled=false, WithWindow=false
try {
  $n = $p.Slides.Count
  $list = if ($Slides) { $Slides.Split(",") | ForEach-Object { [int]$_ } } else { 1..$n }
  foreach ($i in $list) {
    $f = Join-Path $OutDir ("slide-{0:D2}.png" -f $i)
    $p.Slides.Item($i).Export($f, "PNG", 1600, 900)
  }
  Write-Output "exported $($list.Count) of $n"
} finally {
  $p.Close()
  Remove-Item $tmp -ErrorAction SilentlyContinue
}
