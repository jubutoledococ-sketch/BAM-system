$ErrorActionPreference = 'Stop'
Set-Location 'C:\xampp\htdocs\toledo'
$mysql = 'C:\xampp\mysql\bin\mysql.exe'
$outFile = Join-Path (Get-Location) ('inventory_realtime_export_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '_fixed.xlsx')
$tempRoot = Join-Path $env:TEMP ('xlsx_export_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'docProps') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\worksheets') | Out-Null

function Escape-Xml([string]$s) { if ($null -eq $s) { $s = '' }; [System.Security.SecurityElement]::Escape($s) }
function Col-Name([int]$index) { $name=''; while ($index -ge 0) { $name=[char](65+($index%26)) + $name; $index=[math]::Floor($index/26)-1 }; $name }
$sharedMap=@{}; $sharedList=New-Object System.Collections.Generic.List[string]
function Shared-Id([string]$value) { if ($null -eq $value) { $value='' }; if (-not $sharedMap.ContainsKey($value)) { $sharedMap[$value]=$sharedList.Count; $sharedList.Add($value) }; $sharedMap[$value] }
function Get-Data([string]$sql,[string[]]$headers){ $lines=& $mysql --default-character-set=utf8mb4 -u root inventory_db -N -B -e $sql; $rows=@(); foreach($line in $lines){ if($null -eq $line){continue}; $parts=[string]$line -split "`t",-1; $row=@(); for($i=0;$i -lt $headers.Length;$i++){ if($i -lt $parts.Length){$row += [string]$parts[$i]} else {$row += ''} }; $rows += ,$row }; $rows }
function Write-File([string]$path,[string]$content){ [System.IO.File]::WriteAllText($path,$content,[System.Text.UTF8Encoding]::new($false)) }
function Write-Sheet([string]$path,[string]$sheetName,[string[]]$headers,[object[]]$rows){
  $columnCount=[math]::Max($headers.Length,1); $headerRow=4; $dataStartRow=5; $rowCount=$rows.Count+($dataStartRow-1)
  $sb=New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
  [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><tabColor rgb="FF667EEA"/></sheetPr><sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>')
  for($c=0;$c -lt $headers.Length;$c++){ $width=[math]::Max(18,[math]::Min(40,$headers[$c].Length+8)); [void]$sb.Append('<col min="'+($c+1)+'" max="'+($c+1)+'" width="'+$width+'" customWidth="1"/>') }
  [void]$sb.Append('</cols><sheetData>')
  $titleId=Shared-Id $sheetName; $subtitleId=Shared-Id 'Real-time inventory data styled to match the website theme'
  [void]$sb.Append('<row r="1" ht="28" customHeight="1"><c r="A1" t="s" s="4"><v>'+$titleId+'</v></c>')
  for($c=1;$c -lt $columnCount;$c++){ [void]$sb.Append('<c r="'+(Col-Name $c)+'1" s="4"/>') }
  [void]$sb.Append('</row>')
  [void]$sb.Append('<row r="2" ht="22" customHeight="1"><c r="A2" t="s" s="5"><v>'+$subtitleId+'</v></c>')
  for($c=1;$c -lt $columnCount;$c++){ [void]$sb.Append('<c r="'+(Col-Name $c)+'2" s="5"/>') }
  [void]$sb.Append('</row><row r="3" ht="8" customHeight="1">')
  for($c=0;$c -lt $columnCount;$c++){ [void]$sb.Append('<c r="'+(Col-Name $c)+'3" s="6"/>') }
  [void]$sb.Append('</row><row r="4" ht="24" customHeight="1">')
  for($c=0;$c -lt $headers.Length;$c++){ $sid=Shared-Id $headers[$c]; [void]$sb.Append('<c r="'+(Col-Name $c)+'4" t="s" s="1"><v>'+$sid+'</v></c>') }
  [void]$sb.Append('</row>')
  $r=5
  foreach($row in $rows){ [void]$sb.Append('<row r="'+$r+'" ht="21" customHeight="1">'); for($c=0;$c -lt $headers.Length;$c++){ $value=''; if($c -lt $row.Length){ $value=[string]$row[$c] }; $sid=Shared-Id $value; $styleId=if(($r%2)-eq 0){2}else{3}; [void]$sb.Append('<c r="'+(Col-Name $c)+$r+'" t="s" s="'+$styleId+'"><v>'+$sid+'</v></c>') }; [void]$sb.Append('</row>'); $r++ }
  [void]$sb.Append('</sheetData><mergeCells count="2"><mergeCell ref="A1:'+(Col-Name ($columnCount-1))+'1"/><mergeCell ref="A2:'+(Col-Name ($columnCount-1))+'2"/></mergeCells><autoFilter ref="A4:'+(Col-Name ($columnCount-1))+$rowCount+'"/></worksheet>')
  Write-File $path $sb.ToString()
}

$generatedAt=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$worksheets=@(
 @{Name='Dashboard';Headers=@('Field','Value');Rows=@(@('Workbook','Inventory System Real-Time Export'),@('Generated At',$generatedAt))},
 @{Name='Equipment Categories';Headers=@('Category ID','Category Name');Sql="SELECT id, IFNULL(category_name,'') FROM categories ORDER BY category_name ASC"},
 @{Name='Materials';Headers=@('Material ID','Material Name','Supplier','Price','Image Path','Created At');Sql="SELECT material_id, IFNULL(material_name,''), IFNULL(supplier,''), IFNULL(price,''), IFNULL(image,''), IFNULL(created_at,'') FROM materials ORDER BY material_name ASC"},
 @{Name='Equipment';Headers=@('Equipment ID','Equipment Name','Category','Selling Price','Original Price','Stock','Labor Cost','Material Cost','Image Path');Sql="SELECT p.id, IFNULL(p.name,''), IFNULL(c.category_name,''), IFNULL(p.price,''), IFNULL(p.original_price,''), IFNULL(p.stock,''), IFNULL(p.labor_cost,''), IFNULL(p.material_cost,''), IFNULL(p.image,'') FROM products p LEFT JOIN categories c ON c.id = p.category_id ORDER BY p.name ASC"},
 @{Name='Clients';Headers=@('Client ID','Client Name','Contact','Email','Address','Client Type','Payment Terms','Created At');Sql="SELECT id, IFNULL(customer_name,''), IFNULL(contact,''), IFNULL(email,''), IFNULL(address,''), IFNULL(client_type,''), IFNULL(payment_terms,''), IFNULL(created_at,'') FROM customers ORDER BY customer_name ASC"},
 @{Name='Equipment Purchases';Headers=@('Purchase ID','Equipment','Client','Quantity','Original Cost','Discount','Total Cost','Status','Order Date','Expected Delivery','Completed Date','Created By','Notes');Sql="SELECT pu.id, IFNULL(pr.name,''), IFNULL(c.customer_name,'Walk-in'), IFNULL(pu.quantity,''), IFNULL(pu.original_cost,''), IFNULL(pu.discount,''), IFNULL(pu.cost,''), IFNULL(pu.status,''), IFNULL(pu.date,''), IFNULL(pu.expected_delivery,''), IFNULL(pu.completed_date,''), IFNULL(pu.created_by,''), IFNULL(pu.notes,'') FROM purchases pu LEFT JOIN products pr ON pr.id = pu.product_id LEFT JOIN customers c ON c.id = pu.customer_id ORDER BY pu.id DESC"},
 @{Name='Sales Report';Headers=@('Sale ID','Equipment','Client','Purchase ID','Quantity','Total Price','Payment Method','Sale Date');Sql="SELECT s.id, IFNULL(p.name,''), IFNULL(c.customer_name,'Walk-in'), IFNULL(s.purchase_id,''), IFNULL(s.quantity,''), IFNULL(s.total_price,''), IFNULL(pay.payment_method,''), IFNULL(s.date,'') FROM sales s LEFT JOIN products p ON p.id = s.product_id LEFT JOIN purchases pu ON pu.id = s.purchase_id LEFT JOIN customers c ON c.id = pu.customer_id LEFT JOIN payments pay ON pay.sale_id = s.id ORDER BY s.date DESC, s.id DESC"}
)
for($i=1;$i -lt $worksheets.Count;$i++){ $worksheets[$i].Rows = Get-Data $worksheets[$i].Sql $worksheets[$i].Headers }
$worksheets[0].Rows += ,@('Categories',[string]$worksheets[1].Rows.Count)
$worksheets[0].Rows += ,@('Materials',[string]$worksheets[2].Rows.Count)
$worksheets[0].Rows += ,@('Equipment',[string]$worksheets[3].Rows.Count)
$worksheets[0].Rows += ,@('Clients',[string]$worksheets[4].Rows.Count)
$worksheets[0].Rows += ,@('Equipment Purchases',[string]$worksheets[5].Rows.Count)
$worksheets[0].Rows += ,@('Sales Records',[string]$worksheets[6].Rows.Count)

for($i=0;$i -lt $worksheets.Count;$i++){ Write-Sheet (Join-Path $tempRoot ('xl\worksheets\sheet'+($i+1)+'.xml')) $worksheets[$i].Name $worksheets[$i].Headers $worksheets[$i].Rows }

$content='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + ((1..$worksheets.Count|ForEach-Object{'<Override PartName="/xl/worksheets/sheet'+$_+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'}) -join '') + '</Types>'
Write-File (Join-Path $tempRoot '[Content_Types].xml') $content
Write-File (Join-Path $tempRoot '_rels\.rels') '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
Write-File (Join-Path $tempRoot 'docProps\core.xml') ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">'+(Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')+'</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">'+(Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')+'</dcterms:modified></cp:coreProperties>')
$titles=($worksheets|ForEach-Object{'<vt:lpstr>'+(Escape-Xml $_.Name)+'</vt:lpstr>'}) -join ''
Write-File (Join-Path $tempRoot 'docProps\app.xml') ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>'+$worksheets.Count+'</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="'+$worksheets.Count+'" baseType="lpstr">'+$titles+'</vt:vector></TitlesOfParts></Properties>')
$wb=((0..($worksheets.Count-1))|ForEach-Object{'<sheet name="'+(Escape-Xml $worksheets[$_].Name)+'" sheetId="'+($_+1)+'" r:id="rId'+($_+1)+'"/>'}) -join ''
Write-File (Join-Path $tempRoot 'xl\workbook.xml') ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+$wb+'</sheets></workbook>')
$rels=((1..$worksheets.Count|ForEach-Object{'<Relationship Id="rId'+$_+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+$_+'.xml"/>'}) -join '') + '<Relationship Id="rId'+($worksheets.Count+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId'+($worksheets.Count+2)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
Write-File (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+$rels+'</Relationships>')
$styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4"><font><sz val="11"/><color rgb="FF2C3E50"/><name val="Segoe UI"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><sz val="11"/><color rgb="FF334155"/><name val="Segoe UI"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF667EEA"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F7FA"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEAEFFD"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF764BA2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC3CFE2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD4DAE8"/></left><right style="thin"><color rgb="FFD4DAE8"/></right><top style="thin"><color rgb="FFD4DAE8"/></top><bottom style="thin"><color rgb="FFD4DAE8"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
Write-File (Join-Path $tempRoot 'xl\styles.xml') $styles
$shared='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+$sharedList.Count+'" uniqueCount="'+$sharedList.Count+'">' + (($sharedList|ForEach-Object{'<si><t>'+(Escape-Xml $_)+'</t></si>'}) -join '') + '</sst>'
Write-File (Join-Path $tempRoot 'xl\sharedStrings.xml') $shared

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream = [System.IO.File]::Open($outFile, [System.IO.FileMode]::Create)
$zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
$entryMap = @(
    @{ Entry = '[Content_Types].xml'; File = (Join-Path $tempRoot '[Content_Types].xml') },
    @{ Entry = '_rels/.rels'; File = (Join-Path $tempRoot '_rels\.rels') },
    @{ Entry = 'docProps/app.xml'; File = (Join-Path $tempRoot 'docProps\app.xml') },
    @{ Entry = 'docProps/core.xml'; File = (Join-Path $tempRoot 'docProps\core.xml') },
    @{ Entry = 'xl/sharedStrings.xml'; File = (Join-Path $tempRoot 'xl\sharedStrings.xml') },
    @{ Entry = 'xl/styles.xml'; File = (Join-Path $tempRoot 'xl\styles.xml') },
    @{ Entry = 'xl/workbook.xml'; File = (Join-Path $tempRoot 'xl\workbook.xml') },
    @{ Entry = 'xl/_rels/workbook.xml.rels'; File = (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') }
)
foreach($i in 1..$worksheets.Count){
    $entryMap += @{ Entry = ('xl/worksheets/sheet' + $i + '.xml'); File = (Join-Path $tempRoot ('xl\worksheets\sheet' + $i + '.xml')) }
}
foreach($item in $entryMap){
    $entry = $zip.CreateEntry($item.Entry)
    $entryStream = $entry.Open()
    $bytes = [System.IO.File]::ReadAllBytes($item.File)
    $entryStream.Write($bytes,0,$bytes.Length)
    $entryStream.Dispose()
}
$zip.Dispose()
$stream.Dispose()
Remove-Item $tempRoot -Recurse -Force
Write-Output $outFile
