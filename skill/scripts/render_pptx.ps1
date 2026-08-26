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
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-PptxSlideCount {
    $archive = [System.IO.Compression.ZipFile]::OpenRead($inputPath)
    try {
        $count = @($archive.Entries | Where-Object {
            $_.FullName -match '^ppt/slides/slide\d+\.xml$'
        }).Count
        if ($count -lt 1) { throw "No slide XML files found in $inputPath" }
        return $count
    } finally {
        $archive.Dispose()
    }
}

function Clear-PreviousRenderArtifacts {
    if (Test-Path -LiteralPath $pdfPath) {
        [System.IO.File]::Delete($pdfPath)
    }
    Get-ChildItem -LiteralPath $outPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^slide(?:\d+|-\d+)\.png$' } |
        ForEach-Object { [System.IO.File]::Delete($_.FullName) }
}

function Assert-RenderedSlides([int]$ExpectedCount) {
    $slides = @(Get-ChildItem -LiteralPath $outPath -File -Filter 'slide*.png' |
        Where-Object { $_.Name -match '^slide\d+\.png$' })
    if ($slides.Count -ne $ExpectedCount) {
        throw "Expected $ExpectedCount rendered PNGs, found $($slides.Count) in $outPath"
    }
    foreach ($slide in $slides) {
        if ($slide.Length -le 0) { throw "Rendered PNG is empty: $($slide.FullName)" }
    }
}

$sourceSlideCount = Get-PptxSlideCount
Clear-PreviousRenderArtifacts

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

function Resolve-RenderExecutable([string]$Command) {
    $fromPath = Get-Command $Command -ErrorAction SilentlyContinue
    if ($fromPath) { return $fromPath.Source }
    if ($Command -eq "soffice") {
        $candidates = @(
            (Join-Path ${env:ProgramFiles} "LibreOffice\program\soffice.exe"),
            (Join-Path ${env:ProgramFiles(x86)} "LibreOffice\program\soffice.exe"),
            (Join-Path ${env:LOCALAPPDATA} "LibreOffice\program\soffice.exe")
        )
        foreach ($candidate in $candidates) {
            if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
        }
        $uninstallRoots = @(
            "HKLM:\SOFTWARE\LibreOffice\UNO\InstallPath",
            "HKLM:\SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath",
            "HKCU:\SOFTWARE\LibreOffice\UNO\InstallPath"
        )
        foreach ($root in $uninstallRoots) {
            try {
                $installPath = (Get-ItemProperty -Path $root -Name '(default)' -ErrorAction Stop).'(default)'
                $candidate = Join-Path $installPath "soffice.exe"
                if (Test-Path -LiteralPath $candidate) { return $candidate }
            } catch { }
        }
    } elseif ($Command -eq "pdftoppm") {
        $roots = @(${env:ProgramFiles}, ${env:ProgramFiles(x86)}, ${env:LOCALAPPDATA}) |
            Where-Object { $_ }
        foreach ($root in $roots) {
            $packages = Join-Path $root "Microsoft\WinGet\Packages"
            if (-not (Test-Path -LiteralPath $packages)) { continue }
            $candidate = Get-ChildItem -Path $packages -Filter "pdftoppm.exe" -File -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($candidate) { return $candidate.FullName }
        }
    }
    return $null
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
        $soffice = Resolve-RenderExecutable "soffice"
        $pdftoppm = Resolve-RenderExecutable "pdftoppm"
        if (-not $soffice) { throw "LibreOffice (soffice) was not found" }
        if (-not $pdftoppm) { throw "Poppler (pdftoppm) was not found" }
        & $soffice --headless --convert-to pdf --outdir $outPath $inputPath | Out-Null
        if (-not (Test-Path -LiteralPath $pdfPath)) {
            throw "LibreOffice did not create $pdfPath"
        }
        for ($i = 1; $i -le $sourceSlideCount; $i++) {
            $prefix = Join-Path $outPath ("slide{0:D2}" -f $i)
            & $pdftoppm -png -r 96 -f $i -l $i -singlefile $pdfPath $prefix
            if ($LASTEXITCODE -ne 0) { throw "Poppler failed to render slide $i" }
        }
        Write-Output ("EXPORTED via LibreOffice + Poppler -> {0}" -f $outPath)
    } catch {
        Write-Error $_
        exit 1
    }
}

Assert-RenderedSlides -ExpectedCount $sourceSlideCount
