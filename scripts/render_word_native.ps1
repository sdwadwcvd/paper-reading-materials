param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf
)
$ErrorActionPreference = 'Stop'
$sourcePath = (Resolve-Path -LiteralPath $InputDocx).Path
$destinationPath = [System.IO.Path]::GetFullPath($OutputPdf)
if ([string]::Equals($sourcePath, $destinationPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'The output PDF cannot overwrite the source document.'
}
if ([System.IO.Path]::GetExtension($destinationPath) -ne '.pdf') {
    throw 'OutputPdf must have a .pdf extension.'
}
[System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($destinationPath)) | Out-Null
$wordApp = $null
$document = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $document = $wordApp.Documents.Open($sourcePath, $false, $true)
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    $document.ExportAsFixedFormat($destinationPath, 17)
    Write-Output "Rendered pages: $pages"
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    }
    if ($null -ne $wordApp) {
        try { $wordApp.Quit(0) } finally { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
