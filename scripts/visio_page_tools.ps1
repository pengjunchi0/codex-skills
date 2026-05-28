param(
    [Parameter(Mandatory=$true)]
    [string]$VsdxPath,

    [string]$PreviewPath,

    [switch]$Backup,
    [switch]$ExportPreview,
    [switch]$InspectPackage,
    [switch]$CloseOpenDocument
)

$ErrorActionPreference = 'Stop'

function Close-VisioDocument([string]$path) {
    try {
        $visio = [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
        for ($i = $visio.Documents.Count; $i -ge 1; $i--) {
            $doc = $visio.Documents.Item($i)
            if ([string]::Equals($doc.FullName, $path, [StringComparison]::OrdinalIgnoreCase)) {
                $doc.Save()
                $doc.Close()
                Write-Output "Closed open Visio document: $path"
            }
        }
        if ($visio.Documents.Count -eq 0) {
            $visio.Quit()
            Write-Output 'Closed empty Visio instance.'
        }
    } catch {
        Write-Output "No controllable Visio instance found: $($_.Exception.Message)"
    }
}

function Backup-Vsdx([string]$path) {
    $dir = Split-Path -Parent $path
    $stem = [IO.Path]::GetFileNameWithoutExtension($path)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupPath = Join-Path $dir "$stem.backup-$stamp.vsdx"
    Copy-Item -LiteralPath $path -Destination $backupPath
    Write-Output "Backup: $backupPath"
}

function Export-VisioPreview([string]$path, [string]$outPath) {
    if (-not $outPath) {
        $outPath = [IO.Path]::ChangeExtension($path, '.preview.png')
    }
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $true
    $doc = $visio.Documents.Open($path)
    $page = $doc.Pages.Item(1)
    $page.Export($outPath)
    $doc.Close()
    $visio.Quit()
    Write-Output "Preview: $outPath"
}

function Inspect-VsdxPackage([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $media = @($zip.Entries | Where-Object { $_.FullName -like 'visio/media/*' } | Sort-Object FullName)
        $page = $zip.GetEntry('visio/pages/page1.xml')
        $shapeCount = 'unknown'
        if ($page) {
            $reader = [IO.StreamReader]::new($page.Open())
            $xmlText = $reader.ReadToEnd()
            $reader.Close()
            [xml]$xml = $xmlText
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace('v', 'http://schemas.microsoft.com/office/visio/2012/main')
            $shapeCount = $xml.SelectNodes('//v:Shape', $ns).Count
        }
        Write-Output "Shape count: $shapeCount"
        if ($media.Count -eq 0) {
            Write-Output 'Media entries: none'
        } else {
            foreach ($m in $media) {
                Write-Output ("Media: {0} ({1} bytes)" -f $m.FullName, $m.Length)
            }
        }
        $largeMedia = @($media | Where-Object { $_.Length -gt 1000000 -or $_.FullName -match '\.(png|jpg|jpeg)$' })
        Write-Output "Large or raster media entries: $($largeMedia.Count)"
    } finally {
        $zip.Dispose()
    }
}

if ($CloseOpenDocument) { Close-VisioDocument $VsdxPath }
if ($Backup) { Backup-Vsdx $VsdxPath }
if ($InspectPackage) { Inspect-VsdxPackage $VsdxPath }
if ($ExportPreview) { Export-VisioPreview $VsdxPath $PreviewPath }
