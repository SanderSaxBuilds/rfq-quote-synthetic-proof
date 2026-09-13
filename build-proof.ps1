$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $root 'source-rfq.txt'
$pdfPath = Join-Path $root 'sample-rfq.pdf'
$extractPath = Join-Path $root 'extracted-rfq.txt'
$catalogPath = Join-Path $root 'catalog.csv'
$jsonPath = Join-Path $root 'result.json'
$xlsxPath = Join-Path $root 'quote-draft.xlsx'

$sourceLines = Get-Content -LiteralPath $sourcePath

function ConvertTo-PdfLiteral([string]$value) {
  $value = $value -replace '\\', '\\'
  $value = $value -replace '\(', '\('
  $value = $value -replace '\)', '\)'
  return $value
}

$contentLines = @('BT', '/F1 11 Tf', '72 748 Td')
foreach ($line in $sourceLines) {
  $escaped = ConvertTo-PdfLiteral $line
  $contentLines += "($escaped) Tj"
  $contentLines += '0 -16 Td'
}
$contentLines += 'ET'
$content = ($contentLines -join "`n") + "`n"
$contentLength = [Text.Encoding]::ASCII.GetByteCount($content)

$objects = @(
  '<< /Type /Catalog /Pages 2 0 R >>',
  '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
  '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
  '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  "<< /Length $contentLength >>`nstream`n$content" + 'endstream'
)

$pdf = "%PDF-1.4`n"
$offsets = @()
for ($i = 0; $i -lt $objects.Count; $i++) {
  $offsets += [Text.Encoding]::ASCII.GetByteCount($pdf)
  $objectNumber = $i + 1
  $pdf += "$objectNumber 0 obj`n$($objects[$i])`nendobj`n"
}
$xrefStart = [Text.Encoding]::ASCII.GetByteCount($pdf)
$pdf += "xref`n0 6`n0000000000 65535 f `n"
foreach ($offset in $offsets) {
  $pdf += ('{0:D10} 00000 n ' -f $offset) + "`n"
}
$pdf += "trailer`n<< /Size 6 /Root 1 0 R >>`nstartxref`n$xrefStart`n%%EOF`n"
[IO.File]::WriteAllBytes($pdfPath, [Text.Encoding]::ASCII.GetBytes($pdf))

$pdfText = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($pdfPath))
$matches = [regex]::Matches($pdfText, '\((?<text>(?:\\.|[^\\)])*)\)\s*Tj')
$extractedLines = foreach ($match in $matches) {
  $value = $match.Groups['text'].Value
  $value = $value -replace '\\\(', '('
  $value = $value -replace '\\\)', ')'
  $value = $value -replace '\\\\', '\'
  $value
}
$extracted = $extractedLines -join "`r`n"
$extracted | Set-Content -LiteralPath $extractPath -Encoding utf8

$catalog = Import-Csv -LiteralPath $catalogPath
$catalogBySku = @{}
foreach ($row in $catalog) {
  $catalogBySku[$row.sku] = $row
}

$items = @()
foreach ($line in ($extracted -split "`r?`n")) {
  if ($line -match '^(SKU-[0-9]+)\s*\|\s*(.*?)\s*\|\s*Qty\s*([0-9]+)') {
    $sku = $Matches[1]
    $requestedDescription = $Matches[2].Trim()
    $quantity = [int]$Matches[3]
    if ($catalogBySku.ContainsKey($sku)) {
      $approved = $catalogBySku[$sku]
      $unitPrice = [decimal]$approved.unit_price
      $lineTotal = $unitPrice * $quantity
      $items += [pscustomobject]@{
        sku = $sku
        requested_description = $requestedDescription
        approved_description = $approved.description
        quantity = $quantity
        unit_price = $unitPrice
        line_total = $lineTotal
        currency = $approved.currency
        status = 'MATCHED_APPROVED_PRICE'
        note = ''
      }
    }
    else {
      $items += [pscustomobject]@{
        sku = $sku
        requested_description = $requestedDescription
        approved_description = ''
        quantity = $quantity
        unit_price = $null
        line_total = $null
        currency = 'USD'
        status = 'REVIEW_REQUIRED'
        note = 'SKU is not in the approved catalogue. No price guessed.'
      }
    }
  }
}

$result = [pscustomobject]@{
  source_pdf = 'sample-rfq.pdf'
  extracted_text = 'extracted-rfq.txt'
  catalogue = 'catalog.csv'
  policy = 'Exact SKU matches only. Unknown SKUs are unpriced and require review.'
  items = $items
}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding utf8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $workbook = $excel.Workbooks.Add()
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = 'Quote Draft'
  $headers = @('SKU','Requested Description','Approved Description','Qty','Unit Price','Line Total','Currency','Status','Note')
  for ($col = 1; $col -le $headers.Count; $col++) {
    $sheet.Cells.Item(1, $col).Value2 = $headers[$col - 1]
  }
  $rowIndex = 2
  foreach ($item in $items) {
    $sheet.Cells.Item($rowIndex, 1).Value2 = $item.sku
    $sheet.Cells.Item($rowIndex, 2).Value2 = $item.requested_description
    $sheet.Cells.Item($rowIndex, 3).Value2 = $item.approved_description
    $sheet.Cells.Item($rowIndex, 4).Value2 = [string]$item.quantity
    if ($null -ne $item.unit_price) {
      $sheet.Cells.Item($rowIndex, 5).Value2 = $item.unit_price.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
      $sheet.Cells.Item($rowIndex, 6).Value2 = $item.line_total.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
    }
    $sheet.Cells.Item($rowIndex, 7).Value2 = $item.currency
    $sheet.Cells.Item($rowIndex, 8).Value2 = $item.status
    $sheet.Cells.Item($rowIndex, 9).Value2 = $item.note
    $rowIndex++
  }
  $sheet.Range('A1:I1').Font.Bold = $true
  $sheet.Columns.AutoFit() | Out-Null
  $workbook.SaveAs($xlsxPath, 51)
  $workbook.Close($false)
}
finally {
  $excel.Quit()
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}

Write-Output "PDF=$pdfPath"
Write-Output "EXTRACT=$extractPath"
Write-Output "RESULT=$jsonPath"
Write-Output "XLSX=$xlsxPath"
