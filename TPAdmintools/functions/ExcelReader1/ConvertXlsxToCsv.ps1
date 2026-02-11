<# 
ConvertXlsxToCsv.ps1
Convierte un XLSX a CSV usando Excel COM (sin módulos externos).
Requisito: Excel instalado.

Notas:
- No tocamos CutCopyMode (en algunas versiones COM no permite asignación)
- Exporta una única hoja (por índice o nombre)
- Opcional: reescribe delimitador y fuerza UTF-8
#>

param(
  [Parameter(Mandatory=$true)]
  [string]$XlsxPath,

  # 1 = primera hoja, o nombre exacto (string)
  [Parameter(Mandatory=$false)]
  [object]$Sheet = 1,

  # Ruta salida CSV (si no se pasa, mismo nombre/carpeta)
  [Parameter(Mandatory=$false)]
  [string]$CsvPath = "",

  # Delimitador deseado en el CSV (por defecto ;)
  [Parameter(Mandatory=$false)]
  [string]$Delimiter = ";",

  # Fuerza UTF-8 (sin BOM) reescribiendo el archivo final
  [Parameter(Mandatory=$false)]
  [switch]$Utf8
)

function Release-ComObject($obj) {
  if ($null -ne $obj) {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
  }
}

$excel = $null
$workbook = $null
$worksheet = $null
$tempWorkbook = $null
$tempWorksheet = $null
$usedRange = $null
$tmpCsv = $null

try {
  if (-not (Test-Path -LiteralPath $XlsxPath)) {
    throw "No existe el archivo: $XlsxPath"
  }

  if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $CsvPath = [System.IO.Path]::ChangeExtension($XlsxPath, ".csv")
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Open($XlsxPath)

  # Selección de hoja
  if ($Sheet -is [int]) {
    $worksheet = $workbook.Worksheets.Item([int]$Sheet)
  } else {
    $worksheet = $workbook.Worksheets.Item([string]$Sheet)
  }

  # Crear libro temporal y copiar SOLO valores
  $tempWorkbook = $excel.Workbooks.Add()
  $tempWorksheet = $tempWorkbook.Worksheets.Item(1)

  $usedRange = $worksheet.UsedRange

  # Copia valores: asignación directa Value2 = Value2 (sin clipboard)
  $dest = $tempWorksheet.Range("A1").Resize($usedRange.Rows.Count, $usedRange.Columns.Count)
  $dest.Value2 = $usedRange.Value2

  # Exportar CSV base con configuración regional de Excel
  $tmp = [System.IO.Path]::GetTempFileName()
  $tmpCsv = [System.IO.Path]::ChangeExtension($tmp, ".csv")

  # 6 = xlCSV
  $tempWorkbook.SaveAs($tmpCsv, 6)

  # Cerrar todo
  $tempWorkbook.Close($false) | Out-Null
  $workbook.Close($false) | Out-Null
  $excel.Quit() | Out-Null

  # Liberar COM (orden inverso)
  Release-ComObject $dest
  Release-ComObject $usedRange
  Release-ComObject $tempWorksheet
  Release-ComObject $tempWorkbook
  Release-ComObject $worksheet
  Release-ComObject $workbook
  Release-ComObject $excel

  # Reescribir delimitador si Excel sacó , o ; según regional
  $content = Get-Content -LiteralPath $tmpCsv -Raw

  # Detectar delimitador más probable (simple)
  $countComma = ($content.ToCharArray() | Where-Object { $_ -eq ',' }).Count
  $countSemi  = ($content.ToCharArray() | Where-Object { $_ -eq ';' }).Count
  $currentDelim = if ($countSemi -ge $countComma) { ';' } else { ',' }

  if ($currentDelim -ne $Delimiter) {
    $content = $content -replace [Regex]::Escape($currentDelim), $Delimiter
  }

  if ($Utf8) {
    [System.IO.File]::WriteAllText($CsvPath, $content, (New-Object System.Text.UTF8Encoding($false)))
  } else {
    Set-Content -LiteralPath $CsvPath -Value $content
  }

  Remove-Item -LiteralPath $tmpCsv -ErrorAction SilentlyContinue

  Write-Host $CsvPath
  exit 0
}
catch {
  # Intentar cerrar Excel si algo falla
  try { if ($tempWorkbook) { $tempWorkbook.Close($false) | Out-Null } } catch {}
  try { if ($workbook) { $workbook.Close($false) | Out-Null } } catch {}
  try { if ($excel) { $excel.Quit() | Out-Null } } catch {}

  # Liberar COM
  Release-ComObject $usedRange
  Release-ComObject $tempWorksheet
  Release-ComObject $tempWorkbook
  Release-ComObject $worksheet
  Release-ComObject $workbook
  Release-ComObject $excel

  Write-Error $_.Exception.Message
  exit 1
}
finally {
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
