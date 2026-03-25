<?php
require_once 'config.php';
require_once 'db.php';
require_once 'auth.php';

checkLogin();

function fetchRows(mysqli $conn, $sql)
{
    $result = mysqli_query($conn, $sql);

    if (!$result) {
        die('Database error: ' . mysqli_error($conn));
    }

    $rows = [];
    while ($row = mysqli_fetch_assoc($result)) {
        $rows[] = $row;
    }

    return $rows;
}

function rowsOrEmpty(array $rows, array $headers, $message = 'No records found')
{
    if (!empty($rows)) {
        return array_map('array_values', $rows);
    }

    $emptyRow = [$message];
    while (count($emptyRow) < count($headers)) {
        $emptyRow[] = '';
    }

    return [$emptyRow];
}

function xmlValue($value)
{
    return htmlspecialchars((string)$value, ENT_XML1 | ENT_COMPAT, 'UTF-8');
}

function colName($index)
{
    $name = '';
    while ($index >= 0) {
        $name = chr(($index % 26) + 65) . $name;
        $index = intdiv($index, 26) - 1;
    }
    return $name;
}

function worksheetXml($sheetName, array $headers, array $rows, array &$sharedStringMap, array &$sharedStrings)
{
    $columnCount = max(count($headers), 1);
    $headerRow = 4;
    $dataStartRow = 5;
    $rowCount = count($rows) + ($dataStartRow - 1);

    $xml = [];
    $xml[] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $xml[] = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    $xml[] = '<sheetPr><tabColor rgb="FF667EEA"/></sheetPr>';
    $xml[] = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>';
    $xml[] = '<sheetFormatPr defaultRowHeight="15"/>';
    $xml[] = '<cols>';
    foreach ($headers as $index => $header) {
        $width = max(16, min(38, strlen((string)$header) + 6));
        $xml[] = '<col min="' . ($index + 1) . '" max="' . ($index + 1) . '" width="' . $width . '" customWidth="1"/>';
    }
    $xml[] = '</cols>';
    $xml[] = '<sheetData>';

    $titleId = sharedStringId($sheetName, $sharedStringMap, $sharedStrings);
    $subtitleId = sharedStringId('Real-time inventory data styled to match the website theme', $sharedStringMap, $sharedStrings);

    $xml[] = '<row r="1" ht="26" customHeight="1">';
    $xml[] = '<c r="A1" t="s" s="4"><v>' . $titleId . '</v></c>';
    for ($index = 1; $index < $columnCount; $index++) {
        $cellRef = colName($index) . '1';
        $xml[] = '<c r="' . $cellRef . '" s="4"/>';
    }
    $xml[] = '</row>';

    $xml[] = '<row r="2" ht="21" customHeight="1">';
    $xml[] = '<c r="A2" t="s" s="5"><v>' . $subtitleId . '</v></c>';
    for ($index = 1; $index < $columnCount; $index++) {
        $cellRef = colName($index) . '2';
        $xml[] = '<c r="' . $cellRef . '" s="5"/>';
    }
    $xml[] = '</row>';

    $xml[] = '<row r="3" ht="8" customHeight="1">';
    for ($index = 0; $index < $columnCount; $index++) {
        $cellRef = colName($index) . '3';
        $xml[] = '<c r="' . $cellRef . '" s="6"/>';
    }
    $xml[] = '</row>';

    $xml[] = '<row r="' . $headerRow . '" ht="22" customHeight="1">';
    foreach ($headers as $index => $header) {
        $cellRef = colName($index) . $headerRow;
        $stringId = sharedStringId($header, $sharedStringMap, $sharedStrings);
        $xml[] = '<c r="' . $cellRef . '" t="s" s="1"><v>' . $stringId . '</v></c>';
    }
    $xml[] = '</row>';

    $rowNumber = $dataStartRow;
    foreach ($rows as $row) {
        $xml[] = '<row r="' . $rowNumber . '" ht="20" customHeight="1">';
        $colIndex = 0;
        foreach ($row as $value) {
            $cellRef = colName($colIndex) . $rowNumber;
            $stringId = sharedStringId($value, $sharedStringMap, $sharedStrings);
            $styleId = ($rowNumber % 2 === 0) ? '2' : '3';
            $xml[] = '<c r="' . $cellRef . '" t="s" s="' . $styleId . '"><v>' . $stringId . '</v></c>';
            $colIndex++;
        }
        $xml[] = '</row>';
        $rowNumber++;
    }

    $xml[] = '</sheetData>';
    $xml[] = '<mergeCells count="2">';
    $xml[] = '<mergeCell ref="A1:' . colName($columnCount - 1) . '1"/>';
    $xml[] = '<mergeCell ref="A2:' . colName($columnCount - 1) . '2"/>';
    $xml[] = '</mergeCells>';
    $xml[] = '<autoFilter ref="A' . $headerRow . ':' . colName($columnCount - 1) . $rowCount . '"/>';
    $xml[] = '</worksheet>';

    return implode('', $xml);
}

function sharedStringId($value, array &$sharedStringMap, array &$sharedStrings)
{
    $key = (string)($value ?? '');
    if (!array_key_exists($key, $sharedStringMap)) {
        $sharedStringMap[$key] = count($sharedStrings);
        $sharedStrings[] = $key;
    }
    return $sharedStringMap[$key];
}

function sharedStringsXml(array $sharedStrings)
{
    $xml = [];
    $xml[] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $xml[] = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . count($sharedStrings) . '" uniqueCount="' . count($sharedStrings) . '">';
    foreach ($sharedStrings as $string) {
        $xml[] = '<si><t>' . xmlValue($string) . '</t></si>';
    }
    $xml[] = '</sst>';
    return implode('', $xml);
}

function workbookXml(array $sheetNames)
{
    $xml = [];
    $xml[] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $xml[] = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    $xml[] = '<sheets>';
    foreach ($sheetNames as $index => $sheetName) {
        $xml[] = '<sheet name="' . xmlValue($sheetName) . '" sheetId="' . ($index + 1) . '" r:id="rId' . ($index + 1) . '"/>';
    }
    $xml[] = '</sheets>';
    $xml[] = '</workbook>';
    return implode('', $xml);
}

function workbookRelsXml($sheetCount)
{
    $xml = [];
    $xml[] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $xml[] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for ($i = 1; $i <= $sheetCount; $i++) {
        $xml[] = '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
    }
    $xml[] = '<Relationship Id="rId' . ($sheetCount + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
    $xml[] = '<Relationship Id="rId' . ($sheetCount + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
    $xml[] = '</Relationships>';
    return implode('', $xml);
}

function rootRelsXml()
{
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        . '</Relationships>';
}

function contentTypesXml($sheetCount)
{
    $xml = [];
    $xml[] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $xml[] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    $xml[] = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    $xml[] = '<Default Extension="xml" ContentType="application/xml"/>';
    $xml[] = '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    $xml[] = '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    $xml[] = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
    $xml[] = '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
    $xml[] = '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    for ($i = 1; $i <= $sheetCount; $i++) {
        $xml[] = '<Override PartName="/xl/worksheets/sheet' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    }
    $xml[] = '</Types>';
    return implode('', $xml);
}

function stylesXml()
{
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        . '<fonts count="4">'
        . '<font><sz val="11"/><color rgb="FF2C3E50"/><name val="Segoe UI"/></font>'
        . '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>'
        . '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>'
        . '<font><sz val="11"/><color rgb="FF334155"/><name val="Segoe UI"/></font>'
        . '</fonts>'
        . '<fills count="7">'
        . '<fill><patternFill patternType="none"/></fill>'
        . '<fill><patternFill patternType="gray125"/></fill>'
        . '<fill><patternFill patternType="solid"><fgColor rgb="FF667EEA"/><bgColor indexed="64"/></patternFill></fill>'
        . '<fill><patternFill patternType="solid"><fgColor rgb="FFF5F7FA"/><bgColor indexed="64"/></patternFill></fill>'
        . '<fill><patternFill patternType="solid"><fgColor rgb="FFEAEFFD"/><bgColor indexed="64"/></patternFill></fill>'
        . '<fill><patternFill patternType="solid"><fgColor rgb="FF764BA2"/><bgColor indexed="64"/></patternFill></fill>'
        . '<fill><patternFill patternType="solid"><fgColor rgb="FFC3CFE2"/><bgColor indexed="64"/></patternFill></fill>'
        . '</fills>'
        . '<borders count="2">'
        . '<border><left/><right/><top/><bottom/><diagonal/></border>'
        . '<border><left style="thin"><color rgb="FFD4DAE8"/></left><right style="thin"><color rgb="FFD4DAE8"/></right><top style="thin"><color rgb="FFD4DAE8"/></top><bottom style="thin"><color rgb="FFD4DAE8"/></bottom><diagonal/></border>'
        . '</borders>'
        . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        . '<cellXfs count="7">'
        . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        . '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        . '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>'
        . '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>'
        . '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
        . '<xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
        . '<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>'
        . '</cellXfs>'
        . '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        . '</styleSheet>';
}

function coreXml($createdAtIso)
{
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        . '<dc:creator>Codex</dc:creator>'
        . '<cp:lastModifiedBy>Codex</cp:lastModifiedBy>'
        . '<dcterms:created xsi:type="dcterms:W3CDTF">' . xmlValue($createdAtIso) . '</dcterms:created>'
        . '<dcterms:modified xsi:type="dcterms:W3CDTF">' . xmlValue($createdAtIso) . '</dcterms:modified>'
        . '</cp:coreProperties>';
}

function appXml(array $sheetNames)
{
    $titles = '';
    foreach ($sheetNames as $sheetName) {
        $titles .= '<vt:lpstr>' . xmlValue($sheetName) . '</vt:lpstr>';
    }

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        . '<Application>Microsoft Excel</Application>'
        . '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>' . count($sheetNames) . '</vt:i4></vt:variant></vt:vector></HeadingPairs>'
        . '<TitlesOfParts><vt:vector size="' . count($sheetNames) . '" baseType="lpstr">' . $titles . '</vt:vector></TitlesOfParts>'
        . '</Properties>';
}

if (!class_exists('ZipArchive')) {
    die('ZipArchive is not enabled in PHP. Please enable the ZIP extension in XAMPP PHP to generate .xlsx files.');
}

$generatedAt = date('Y-m-d H:i:s');
$generatedAtIso = gmdate('Y-m-d\TH:i:s\Z');

$categories = fetchRows($conn, "
    SELECT id AS `Category ID`, category_name AS `Category Name`
    FROM categories
    ORDER BY category_name ASC
");

$materials = fetchRows($conn, "
    SELECT
        material_name AS `Material Name`,
        supplier AS `Supplier`,
        price AS `Price`,
        created_at AS `Created At`
    FROM materials
    ORDER BY material_name ASC
");

$equipment = fetchRows($conn, "
    SELECT
        p.name AS `Equipment Name`,
        COALESCE(c.category_name, '') AS `Category`,
        p.price AS `Selling Price`,
        p.stock AS `Stock`
    FROM products p
    LEFT JOIN categories c ON c.id = p.category_id
    ORDER BY p.name ASC
");

$clients = fetchRows($conn, "
    SELECT
        customer_name AS `Client Name`,
        contact AS `Contact`,
        email AS `Email`,
        address AS `Address`,
        client_type AS `Client Type`,
        payment_terms AS `Payment Terms`,
        created_at AS `Created At`
    FROM customers
    ORDER BY customer_name ASC
");

$sales = fetchRows($conn, "
    SELECT
        p.name AS `Product`,
        s.quantity AS `Qty`,
        s.total_price AS `Total`,
        s.date AS `Sale Date`
    FROM sales s
    JOIN products p ON p.id = s.product_id
    ORDER BY s.date ASC, s.id ASC
");

$dashboard = [
    ['Workbook', 'Inventory System Real-Time Export'],
    ['Generated At', $generatedAt],
    ['Categories', count($categories)],
    ['Materials', count($materials)],
    ['Equipment', count($equipment)],
    ['Clients', count($clients)],
    ['Sales Records', count($sales)],
];

$worksheets = [
    'Dashboard' => [
        'headers' => ['Field', 'Value'],
        'rows' => $dashboard,
    ],
    'Equipment Categories' => [
        'headers' => array_keys($categories[0] ?? ['Category ID' => '', 'Category Name' => '']),
        'rows' => rowsOrEmpty($categories, array_keys($categories[0] ?? ['Category ID' => '', 'Category Name' => ''])),
    ],
    'Materials' => [
        'headers' => array_keys($materials[0] ?? ['Material Name' => '', 'Supplier' => '', 'Price' => '', 'Created At' => '']),
        'rows' => rowsOrEmpty($materials, array_keys($materials[0] ?? ['Material Name' => '', 'Supplier' => '', 'Price' => '', 'Created At' => ''])),
    ],
    'Equipment' => [
        'headers' => array_keys($equipment[0] ?? ['Equipment Name' => '', 'Category' => '', 'Selling Price' => '', 'Stock' => '']),
        'rows' => rowsOrEmpty($equipment, array_keys($equipment[0] ?? ['Equipment Name' => '', 'Category' => '', 'Selling Price' => '', 'Stock' => ''])),
    ],
    'Clients' => [
        'headers' => array_keys($clients[0] ?? ['Client Name' => '', 'Contact' => '', 'Email' => '', 'Address' => '', 'Client Type' => '', 'Payment Terms' => '', 'Created At' => '']),
        'rows' => rowsOrEmpty($clients, array_keys($clients[0] ?? ['Client Name' => '', 'Contact' => '', 'Email' => '', 'Address' => '', 'Client Type' => '', 'Payment Terms' => '', 'Created At' => ''])),
    ],
    'Sales Report' => [
        'headers' => array_keys($sales[0] ?? ['Product' => '', 'Qty' => '', 'Total' => '', 'Sale Date' => '']),
        'rows' => rowsOrEmpty($sales, array_keys($sales[0] ?? ['Product' => '', 'Qty' => '', 'Total' => '', 'Sale Date' => ''])),
    ],
];

$sharedStringMap = [];
$sharedStrings = [];
$sheetXmlFiles = [];
$sheetNames = array_keys($worksheets);

$sheetIndex = 1;
foreach ($worksheets as $sheet) {
    $sheetXmlFiles['xl/worksheets/sheet' . $sheetIndex . '.xml'] = worksheetXml(
        $sheetNames[$sheetIndex - 1],
        $sheet['headers'],
        $sheet['rows'],
        $sharedStringMap,
        $sharedStrings
    );
    $sheetIndex++;
}

$tempFile = tempnam(sys_get_temp_dir(), 'xlsx_export_');
$zip = new ZipArchive();

if ($zip->open($tempFile, ZipArchive::OVERWRITE) !== true) {
    die('Unable to create Excel file.');
}

$zip->addFromString('[Content_Types].xml', contentTypesXml(count($worksheets)));
$zip->addFromString('_rels/.rels', rootRelsXml());
$zip->addFromString('docProps/core.xml', coreXml($generatedAtIso));
$zip->addFromString('docProps/app.xml', appXml($sheetNames));
$zip->addFromString('xl/workbook.xml', workbookXml($sheetNames));
$zip->addFromString('xl/_rels/workbook.xml.rels', workbookRelsXml(count($worksheets)));
$zip->addFromString('xl/styles.xml', stylesXml());
$zip->addFromString('xl/sharedStrings.xml', sharedStringsXml($sharedStrings));

foreach ($sheetXmlFiles as $path => $xml) {
    $zip->addFromString($path, $xml);
}

$zip->close();

$filename = 'inventory_realtime_export_' . date('Ymd_His') . '.xlsx';

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Length: ' . filesize($tempFile));
header('Pragma: no-cache');
header('Expires: 0');

readfile($tempFile);
unlink($tempFile);
exit;
