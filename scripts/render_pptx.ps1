param([string]$InFile)
$ErrorActionPreference = "Stop"
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    Start-Sleep -Milliseconds 1200
    $pres = $ppt.Presentations.Open($InFile, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InFile)
    $outDir = Join-Path (Split-Path $InFile) ("render_" + $base)
    if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $n = $pres.Slides.Count
    for ($i = 1; $i -le $n; $i++) {
        $out = Join-Path $outDir ("slide{0}.png" -f $i)
        $pres.Slides.Item($i).Export($out, "PNG", 1280, 720)
    }
    Write-Output ("EXPORTED {0} slides to {1}" -f $n, $outDir)
    $pres.Close()
} catch {
    Write-Output ("ERROR: " + $_.Exception.Message)
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
