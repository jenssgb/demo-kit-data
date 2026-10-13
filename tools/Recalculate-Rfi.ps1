# Recalculate only the four new RFI masters in an isolated Excel instance.
# Run build-cowork-rfi.py --prepare first, then this, --copies and Update-Manifest.ps1 -All.
$ErrorActionPreference = 'Stop'
$root = Join-Path (Split-Path $PSScriptRoot -Parent) 'cowork-rfi-rfp\templates'
$paths = foreach ($language in 'en', 'de') {
    foreach ($variant in 'Full', 'Compact') {
        $path = Join-Path $root "$language\RFI_${variant}_Master.xlsx"
        if (-not (Test-Path -LiteralPath $path)) { throw "Missing master: $path" }
        $path
    }
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
try {
    foreach ($path in $paths) {
        $book = $excel.Workbooks.Open($path, 0, $false)
        try {
            $excel.CalculateFullRebuild()
            $tco = $book.Worksheets.Item('TCO')
            if ($tco.Range('B3').Value2 -ne 172800) { throw "$path : incorrect managed-service subtotal" }
            if ($tco.Range('B4').Text -ne '') { throw "$path : incomplete net TCO must be blank" }
            if ($tco.Range('B6').Text -ne 'PARTIAL') { throw "$path : wrong net coverage" }
            $book.Save()
            Write-Output ("Recalculated: {0}" -f (Split-Path $path -Leaf))
        } finally {
            $book.Close($false)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
        }
    }
} finally {
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}
