param(
    [Parameter(Mandatory = $true)][string]$InFile,
    [string]$OutDir = ""
)

$script = Join-Path (Split-Path -Parent $PSScriptRoot) "skill\scripts\render_pptx.ps1"
& $script -InFile $InFile -OutDir $OutDir
exit $LASTEXITCODE
