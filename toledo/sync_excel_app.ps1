param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$mysql = 'C:\xampp\mysql\bin\mysql.exe'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-Data([string]$sql, [string[]]$headers) {
    $lines = & $mysql --default-character-set=utf8mb4 -u root inventory_db -N -B -e $sql
    $rows = @()
    foreach ($line in $lines) {
        if ($null -eq $line) { continue }
        $parts = [string]$line -split "`t", -1
        $row = @()
        for ($i = 0; $i -lt $headers.Length; $i++) {
            if ($i -lt $parts.Length) { $row += [string]$parts[$i] } else { $row += '' }
        }
        $rows += ,$row
    }
    return ,$rows
}

function Rows-OrEmpty([object[]]$rows, [int]$colCount) {
    if ($rows.Count -gt 0) { return ,$rows }
    $row = @('No records found')
    while ($row.Count -lt $colCount) { $row += '' }
    return ,$row
}

function Normalize-Rows([object[]]$rows, [int]$colCount) {
    $normalized = @()
    foreach ($row in $rows) {
        $newRow = @()
        if ($row -is [System.Array]) {
            for ($i = 0; $i -lt $colCount; $i++) {
                if ($i -lt $row.Length) {
                    $newRow += [string]$row[$i]
                } else {
                    $newRow += ''
                }
            }
        } else {
            $newRow += [string]$row
            while ($newRow.Count -lt $colCount) {
                $newRow += ''
            }
        }
        $normalized += ,$newRow
    }
    return ,$normalized
}

$categoryHeaders = @('Category ID','Category Name')
$materialHeaders = @('Material Name','Supplier','Price','Created At')
$equipmentHeaders = @('Equipment Name','Category','Selling Price','Stock')
$clientHeaders = @('Client Name','Contact','Email','Address','Client Type','Payment Terms','Created At')
$salesHeaders = @('Product','Qty','Total','Date')

$categories = Get-Data "SELECT id, IFNULL(category_name,'') FROM categories ORDER BY id ASC" $categoryHeaders
$materials = Get-Data "SELECT IFNULL(material_name,''), IFNULL(supplier,''), IFNULL(price,''), IFNULL(created_at,'') FROM materials ORDER BY material_name ASC" $materialHeaders
$equipment = Get-Data "SELECT IFNULL(p.name,''), IFNULL(c.category_name,''), IFNULL(p.price,''), IFNULL(p.stock,'') FROM products p LEFT JOIN categories c ON c.id = p.category_id ORDER BY p.name ASC" $equipmentHeaders
$clients = Get-Data "SELECT IFNULL(customer_name,''), IFNULL(contact,''), IFNULL(email,''), IFNULL(address,''), IFNULL(client_type,''), IFNULL(payment_terms,''), IFNULL(created_at,'') FROM customers ORDER BY customer_name ASC" $clientHeaders
$sales = Get-Data "SELECT IFNULL(p.name,''), IFNULL(s.quantity,''), IFNULL(s.total_price,''), IFNULL(s.date,'') FROM sales s JOIN products p ON p.id = s.product_id ORDER BY s.date ASC, s.id ASC" $salesHeaders

$salesTotalQty = 0.0
$salesTotalAmount = 0.0
foreach ($saleRow in $sales) {
    if ($saleRow.Length -ge 3) {
        $salesTotalQty += [double]($saleRow[1] -as [double])
        $salesTotalAmount += [double]($saleRow[2] -as [double])
    }
}

$salesRows = Rows-OrEmpty $sales $salesHeaders.Count
if ($sales.Count -gt 0) {
    $salesRows += ,@('TOTAL', [string][int]$salesTotalQty, [string]([Math]::Round($salesTotalAmount, 2)), '')
}

$dashboardRows = @(
    @('Workbook', 'Inventory System Real-Time Export'),
    @('Generated At', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
    @('Categories', [string]$categories.Count),
    @('Materials', [string]$materials.Count),
    @('Equipment', [string]$equipment.Count),
    @('Clients', [string]$clients.Count),
    @('Sales Records', [string]$sales.Count)
)

$worksheets = @(
    @{ Name = 'Dashboard'; Headers = @('Field','Value'); Rows = $dashboardRows },
    @{ Name = 'Equipment Categories'; Headers = $categoryHeaders; Rows = (Rows-OrEmpty $categories $categoryHeaders.Count) },
    @{ Name = 'Materials'; Headers = $materialHeaders; Rows = (Rows-OrEmpty $materials $materialHeaders.Count) },
    @{ Name = 'Equipment'; Headers = $equipmentHeaders; Rows = (Rows-OrEmpty $equipment $equipmentHeaders.Count) },
    @{ Name = 'Clients'; Headers = $clientHeaders; Rows = (Rows-OrEmpty $clients $clientHeaders.Count) },
    @{ Name = 'Sales Report'; Headers = $salesHeaders; Rows = $salesRows }
)

for ($i = 0; $i -lt $worksheets.Count; $i++) {
    $worksheets[$i].Rows = Normalize-Rows $worksheets[$i].Rows $worksheets[$i].Headers.Count
}

$excel = $null
$workbook = $null
$createdExcel = $false

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($wb in $excel.Workbooks) {
            if ($wb.FullName -eq $OutputPath) {
                $workbook = $wb
                break
            }
        }
    } catch {
        $excel = $null
    }

    if ($excel -eq $null) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $createdExcel = $true
    }

    if ($workbook -eq $null) {
        if (Test-Path $OutputPath) {
            $workbook = $excel.Workbooks.Open($OutputPath)
        } else {
            $workbook = $excel.Workbooks.Add()
        }
    }

    while ($workbook.Worksheets.Count -lt $worksheets.Count) { $null = $workbook.Worksheets.Add() }
    while ($workbook.Worksheets.Count -gt $worksheets.Count) { $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete() }

    $headerBg = 0x3D5A80
    $subBg = 0xEE6C4D
    $bandBg = 0xE0FBFC
    $rowLight = 0xF7FAFC
    $rowAlt = 0xEAF4F4
    $sheetBg = 0xF4F7F8
    $headerText = 0xFFFFFF
    $bodyText = 0x243B53
    $border = 0xC9D6DF
    $tabColor = 0xEE6C4D

    for ($i = 0; $i -lt $worksheets.Count; $i++) {
        $sheet = $workbook.Worksheets.Item($i + 1)
        $sheet.Name = $worksheets[$i].Name

        $headers = $worksheets[$i].Headers
        $rows = $worksheets[$i].Rows
        $colCount = $headers.Count
        $rowCount = $rows.Count

        if ($sheet.FilterMode) {
            try { $sheet.ShowAllData() | Out-Null } catch {}
        }
        $sheet.AutoFilterMode = $false
        $sheet.Cells.EntireRow.Hidden = $false
        $sheet.Cells.EntireColumn.Hidden = $false
        $sheet.Cells.Clear() | Out-Null
        $sheet.Cells.Interior.Color = $sheetBg

        $titleRange = $sheet.Range($sheet.Cells(1,1), $sheet.Cells(1,$colCount))
        $titleRange.Merge()
        $titleRange.Value2 = $worksheets[$i].Name
        $titleRange.Interior.Color = $headerBg
        $titleRange.Font.Name = 'Segoe UI'
        $titleRange.Font.Color = $headerText
        $titleRange.Font.Bold = $true
        $titleRange.Font.Size = 16
        $titleRange.RowHeight = 28

        $subRange = $sheet.Range($sheet.Cells(2,1), $sheet.Cells(2,$colCount))
        $subRange.Merge()
        $subRange.Value2 = 'Automatically synced from the website database'
        $subRange.Interior.Color = $subBg
        $subRange.Font.Name = 'Segoe UI'
        $subRange.Font.Color = $headerText
        $subRange.Font.Size = 11
        $subRange.RowHeight = 22

        $bandRange = $sheet.Range($sheet.Cells(3,1), $sheet.Cells(3,$colCount))
        $bandRange.Interior.Color = $bandBg
        $bandRange.RowHeight = 8

        for ($c = 0; $c -lt $colCount; $c++) {
            $sheet.Cells.Item(4, $c + 1).Value2 = [string]$headers[$c]
        }

        $headerRange = $sheet.Range($sheet.Cells(4,1), $sheet.Cells(4,$colCount))
        $headerRange.Interior.Color = $headerBg
        $headerRange.Font.Name = 'Segoe UI'
        $headerRange.Font.Color = $headerText
        $headerRange.Font.Bold = $true
        $headerRange.HorizontalAlignment = -4108
        $headerRange.VerticalAlignment = -4108
        $headerRange.Borders.Color = $border
        $headerRange.RowHeight = 24

        for ($r = 0; $r -lt $rowCount; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $sheet.Cells.Item($r + 5, $c + 1).Value2 = [string]$rows[$r][$c]
            }

            $rowRange = $sheet.Range($sheet.Cells($r + 5,1), $sheet.Cells($r + 5,$colCount))
            if ((($r + 5) % 2) -eq 0) {
                $rowRange.Interior.Color = $rowLight
            } else {
                $rowRange.Interior.Color = $rowAlt
            }
            $rowRange.Font.Name = 'Segoe UI'
            $rowRange.Font.Size = 11
            $rowRange.Font.Color = $bodyText
            $rowRange.Borders.Color = $border
            $rowRange.RowHeight = 20
        }

        $endRow = [Math]::Max(5, $rowCount + 4)
        $fullRange = $sheet.Range($sheet.Cells(1,1), $sheet.Cells($endRow,$colCount))
        $fullRange.Borders.Color = $border
        $dataRange = $sheet.Range($sheet.Cells(4,1), $sheet.Cells($endRow,$colCount))
        $dataRange.Columns.AutoFit() | Out-Null
        $dataRange.AutoFilter() | Out-Null

        if ($worksheets[$i].Name -eq 'Sales Report' -and $rowCount -gt 0) {
            $firstDataRow = 5
            $lastDataRow = $endRow
            $sheet.Range($sheet.Cells($firstDataRow, 2), $sheet.Cells($lastDataRow, 2)).NumberFormat = '0'
            $sheet.Range($sheet.Cells($firstDataRow, 3), $sheet.Cells($lastDataRow, 3)).NumberFormat = '#,##0.00'
            if ($rowCount -gt 1) {
                $sheet.Range($sheet.Cells($firstDataRow, 4), $sheet.Cells($lastDataRow - 1, 4)).NumberFormat = 'mmm dd, yyyy'
            }
        }

        $sheet.Activate() | Out-Null
        $excel.ActiveWindow.SplitRow = 4
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.FreezePanes = $true
        $sheet.Range('A5').Select() | Out-Null
        $excel.ActiveWindow.ScrollRow = 1
        $excel.ActiveWindow.ScrollColumn = 1
        $sheet.Tab.Color = $tabColor
    }

    $workbook.Worksheets.Item(1).Activate() | Out-Null
    if (Test-Path $OutputPath) {
        $workbook.Save()
    } else {
        $workbook.SaveAs($OutputPath, 51)
    }
    if ($createdExcel) {
        $workbook.Close($true)
    }
}
finally {
    if ($workbook -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($createdExcel -and $excel -ne $null) {
        $excel.Quit()
    }
    if ($excel -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
