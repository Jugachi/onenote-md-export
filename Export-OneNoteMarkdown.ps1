<#
.SYNOPSIS
  Export local OneNote notebooks to a Markdown folder tree.

.DESCRIPTION
  Thin wrapper around onenote-md.exe so you can call it from PowerShell.
  Builds the executable on first use if it is missing.

.EXAMPLE
  .\Export-OneNoteMarkdown.ps1 -List
.EXAMPLE
  .\Export-OneNoteMarkdown.ps1 -OutputDir D:\notes -Notebook KLG
#>
[CmdletBinding()]
param(
    [string]   $OutputDir   = "onenote-export",
    [string]   $Notebook,
    [string]   $Section,
    [switch]   $List,
    [switch]   $NoImages,
    [switch]   $SkipExisting,
    [switch]   $DryRun
)

$ErrorActionPreference = 'Stop'

$exe = Join-Path $PSScriptRoot 'bin\onenote-md.exe'
if (-not (Test-Path $exe)) {
    Write-Host "Building onenote-md.exe ..." -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'build.cmd')
    if ($LASTEXITCODE -ne 0) { throw "build.cmd failed" }
}

$args = @($OutputDir)
if ($Notebook)      { $args += @('--notebook', $Notebook) }
if ($Section)       { $args += @('--section',  $Section)  }
if ($List)          { $args += '--list' }
if ($NoImages)      { $args += '--no-images' }
if ($SkipExisting)  { $args += '--skip-existing' }
if ($DryRun)        { $args += '--dry-run' }

& $exe @args
exit $LASTEXITCODE
