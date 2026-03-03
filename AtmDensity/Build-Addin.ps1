<#
.SYNOPSIS
    Creates AtmDensity.xlam (Excel Add-in) from NRLMSISE00.bas

.DESCRIPTION
    Launches Excel, imports the VBA module, and saves as .xlam.
    Run this script once; afterwards install the .xlam in Excel.

.NOTES
    Requires Windows + Excel installed.
    Run in PowerShell as a normal user (not administrator).

.EXAMPLE
    .\Build-Addin.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$basFile     = Join-Path $scriptDir 'NRLMSISE00.bas'
$xlamPath    = Join-Path $scriptDir 'AtmDensity.xlam'

if (-not (Test-Path $basFile)) {
    Write-Error "NRLMSISE00.bas not found at: $basFile"
    exit 1
}

Write-Host "Opening Excel..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    # Create a new workbook
    $wb = $xl.Workbooks.Add()

    # Allow VBA macro access (required for programmatic import)
    $xl.Application.VBE.ActiveVBProject | Out-Null

    # Import the .bas module
    Write-Host "Importing NRLMSISE00.bas..."
    $wb.VBProject.VBComponents.Import($basFile) | Out-Null

    # Add a sheet with usage instructions
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Usage"
    $ws.Cells.Item(1,1).Value2 = "NRLMSISE-00 Atmospheric Density Add-in"
    $ws.Cells.Item(2,1).Value2 = "Functions available after installing this add-in:"
    $ws.Cells.Item(4,1).Value2 = "=ATMDENSITY(alt_km, f107, [f107a], [Ap], [doy], [lat], [lon], [lst])"
    $ws.Cells.Item(5,1).Value2 = "  Returns total atmospheric density in kg/m^3"
    $ws.Cells.Item(7,1).Value2 = "=ATMTEMPERATURE(alt_km, f107, [f107a], [Ap], [doy], [lat], [lon], [lst])"
    $ws.Cells.Item(8,1).Value2 = "  Returns exospheric temperature in K"
    $ws.Cells.Item(10,1).Value2 = "=ATMDENSITY_SPECIES(alt_km, f107, species, [f107a], [Ap], [doy], [lat], [lon], [lst])"
    $ws.Cells.Item(11,1).Value2 = "  species: He | O | N2 | O2 | Ar | Total | H | N | AnoO"
    $ws.Cells.Item(13,1).Value2 = "Parameters:"
    $ws.Cells.Item(14,1).Value2 = "  alt_km  - Geodetic altitude (km)"
    $ws.Cells.Item(15,1).Value2 = "  f107    - Daily F10.7 solar flux index"
    $ws.Cells.Item(16,1).Value2 = "  f107a   - 81-day average F10.7 (default = f107)"
    $ws.Cells.Item(17,1).Value2 = "  Ap      - Geomagnetic Ap index (default 4 = quiet)"
    $ws.Cells.Item(18,1).Value2 = "  doy     - Day of year 1-366 (default 172)"
    $ws.Cells.Item(19,1).Value2 = "  lat     - Geodetic latitude degrees (default 45)"
    $ws.Cells.Item(20,1).Value2 = "  lon     - Longitude degrees (default 0)"
    $ws.Cells.Item(21,1).Value2 = "  lst     - Local solar time hours (default 12)"
    $ws.Cells.Item(23,1).Value2 = "Example: =ATMDENSITY(400, 150) -> density at 400 km for F10.7=150"
    $ws.Columns.AutoFit()

    # Save as .xlam
    Write-Host "Saving as $xlamPath ..."
    $wb.SaveAs($xlamPath, 55)   # 55 = xlOpenXMLAddIn (.xlam)
    $wb.Close($false)

    Write-Host ""
    Write-Host "SUCCESS: $xlamPath created."
    Write-Host ""
    Write-Host "To install in Excel:"
    Write-Host "  1. Open Excel > File > Options > Add-ins"
    Write-Host "  2. At the bottom, Manage: Excel Add-ins > Go"
    Write-Host "  3. Click Browse and select: $xlamPath"
    Write-Host "  4. Tick the checkbox and click OK"
    Write-Host ""
    Write-Host "Then use =ATMDENSITY(400, 150) in any cell."

} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
