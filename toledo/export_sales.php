<?php

require_once "db.php"; // Make sure this file sets up $conn correctly

$start = $_GET['start'] ?? date('Y-m-01');
$end   = $_GET['end'] ?? date('Y-m-d');

$sales = mysqli_query($conn, "
    SELECT p.name, s.quantity, s.total_price, s.date
    FROM sales s
    JOIN products p ON s.product_id = p.id
    WHERE s.date BETWEEN '$start' AND '$end'
    ORDER BY s.date ASC
");

if (!$sales) {
    die('Database error: ' . mysqli_error($conn));
}

// Set headers for CSV download
header('Content-Type: text/csv; charset=utf-8');
header('Content-Disposition: attachment; filename=sales_report_' . $start . '_to_' . $end . '.csv');

// Open the output stream
$output = fopen('php://output', 'w');

// Output the column headings
fputcsv($output, ['Product', 'Quantity', 'Total', 'Date']);

// Output the rows
while ($row = mysqli_fetch_assoc($sales)) {
    fputcsv($output, [
        $row['name'],
        $row['quantity'],
        number_format($row['total_price'], 2),
        date('Y-m-d', strtotime($row['date']))
    ]);
}

fclose($output);
exit;
