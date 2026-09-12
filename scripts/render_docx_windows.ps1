param(
    [Parameter(Mandatory = $true)][string]$ArchiveDir,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [string]$PdfToPpmPath = ""
)

$ErrorActionPreference = 'Stop'
$archive = (Resolve-Path -LiteralPath $ArchiveDir).Path
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}
$qaRoot = (Resolve-Path -LiteralPath $OutputDir).Path

if (-not $PdfToPpmPath) {
    $command = Get-Command pdftoppm -ErrorAction SilentlyContinue
    if ($command) { $PdfToPpmPath = $command.Source }
}
if (-not $PdfToPpmPath -or -not (Test-Path -LiteralPath $PdfToPpmPath)) {
    throw 'pdftoppm was not found. Pass -PdfToPpmPath from load_workspace_dependencies.'
}

$documents = Get-ChildItem -LiteralPath $archive -Recurse -Filter '*.docx' -File | Sort-Object FullName
if ($documents.Count -eq 0) { throw "No DOC files found under $archive" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($source in $documents) {
        $folderName = $source.Directory.Name
        $itemOut = Join-Path $qaRoot $folderName
        New-Item -ItemType Directory -Path $itemOut -Force | Out-Null
        Get-ChildItem -LiteralPath $itemOut -File -ErrorAction SilentlyContinue | Remove-Item -Force
        $pdf = Join-Path $itemOut ($source.BaseName + '.pdf')
        $document = $word.Documents.Open($source.FullName, $false, $true)
        try { $document.ExportAsFixedFormat($pdf, 17) }
        finally { $document.Close($false) }
        & $PdfToPpmPath -png -r 120 $pdf (Join-Path $itemOut 'page')
        if ($LASTEXITCODE -ne 0) { throw "pdftoppm failed for $pdf" }
        $pages = (Get-ChildItem -LiteralPath $itemOut -Filter 'page-*.png' -File).Count
        if ($pages -lt 1) { throw "No rendered pages for $($source.FullName)" }
        Write-Output "$folderName`t$pages`t$pdf"
    }
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

