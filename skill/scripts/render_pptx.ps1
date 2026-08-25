param(
    [Parameter(Mandatory = $true)][string]$InFile,
    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"
$inputPath = (Resolve-Path -LiteralPath $InFile).Path
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path (Split-Path -Parent $inputPath) ("render_" + [IO.Path]::GetFileNameWithoutExtension($inputPath))
}
$outPath = [IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $outPath | Out-Null
$pdfPath = Join-Path $outPath ([IO.Path]::GetFileNameWithoutExtension($inputPath) + ".pdf")

function Export-WithPowerPoint {
    $ppt = $null
    $pres = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        # Do not force Visible here: some Office builds reject hiding the
        # application window when launched through COM.
        $pres = $ppt.Presentations.Open($inputPath, $false, $false, $false)
        # SaveAs with ppSaveAsPDF (32) is more compatible with PowerPoint COM
        # than ExportAsFixedFormat on installations with strict COM binding.
        $pres.SaveAs($pdfPath, 32)
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $png = Join-Path $outPath ("slide{0:D2}.png" -f $i)
            $pres.Slides.Item($i).Export($png, "PNG", 1280, 720)
        }
        Write-Output ("EXPORTED via PowerPoint: {0} slides -> {1}" -f $pres.Slides.Count, $outPath)
        return $true
    } finally {
        if ($pres) { $pres.Close() }
        if ($ppt) { $ppt.Quit() }
        if ($pres) { [Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }
        if ($ppt) { [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    }
}

$rendered = $false
try {
    [void](Export-WithPowerPoint)
    $rendered = $true
} catch {
    Write-Warning ("PowerPoint COM unavailable or failed: " + $_.Exception.Message)
}

if (-not $rendered) {
    try {
        $soffice = Get-Command soffice -ErrorAction SilentlyContinue
        $pdftoppm = Get-Command pdftoppm -ErrorAction SilentlyContinue
        if (-not $soffice) { throw "LibreOffice (soffice) was not found" }
        if (-not $pdftoppm) { throw "Poppler (pdftoppm) was not found" }
        & $soffice.Source --headless --convert-to pdf --outdir $outPath $inputPath | Out-Null
        if (-not (Test-Path -LiteralPath $pdfPath)) {
            throw "LibreOffice did not create $pdfPath"
        }
        $prefix = Join-Path $outPath "slide"
        & $pdftoppm.Source -png -r 96 $pdfPath $prefix
        Write-Output ("EXPORTED via LibreOffice + Poppler -> {0}" -f $outPath)
    } catch {
        Write-Error $_
        exit 1
    }
}
