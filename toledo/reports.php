<?php
require_once 'header.php';

$start = $_GET['start'] ?? date('Y-m-01');
$end = $_GET['end'] ?? date('Y-m-d');

// Summary metrics
$report = mysqli_fetch_assoc(mysqli_query($conn,
    "SELECT 
        COUNT(*) AS transactions,
        IFNULL(SUM(quantity),0) AS items,
        IFNULL(SUM(total_price),0) AS income,
        IFNULL(AVG(total_price),0) AS avg_sale
     FROM sales
     WHERE date BETWEEN '$start' AND '$end'"
));

// Top-selling product
$topProduct = mysqli_fetch_assoc(mysqli_query($conn,
    "SELECT p.name, SUM(s.quantity) AS total_qty
     FROM sales s
     JOIN products p ON s.product_id = p.id
     WHERE s.date BETWEEN '$start' AND '$end'
     GROUP BY s.product_id
     ORDER BY total_qty DESC
     LIMIT 1"
));

// Fetch sales data
$sales = mysqli_query($conn, "
    SELECT s.id, p.name, s.quantity, s.total_price, s.date
    FROM sales s
    JOIN products p ON s.product_id = p.id
    WHERE s.date BETWEEN '$start' AND '$end'
    ORDER BY s.date ASC
");

?>

<h1>Sales Report</h1>

<form method="get">
    <label>From</label>
    <input type="date" name="start" value="<?= htmlspecialchars($start) ?>">

    <label>To</label>
    <input type="date" name="end" value="<?= htmlspecialchars($end) ?>">

    <button type="submit">Filter</button>
</form>

<div class="grid">
    <div class="card">
        <h3>Total Transactions</h3>
        <p><?= (int)$report['transactions'] ?></p>
    </div>
    <div class="card">
        <h3>Total Items Sold</h3>
        <p><?= (int)$report['items'] ?></p>
    </div>
    <div class="card">
        <h3>Total Income</h3>
        <p>₱<?= number_format($report['income'], 2) ?></p>
    </div>
    <div class="card">
        <h3>Average Sale</h3>
        <p>₱<?= number_format($report['avg_sale'], 2) ?></p>
    </div>
    <div class="card">
        <h3>Top Selling Product</h3>
        <p>
            <?= htmlspecialchars($topProduct['name'] ?? 'N/A') ?> 
            (<?= $topProduct['total_qty'] ?? 0 ?> sold)
        </p>
    </div>
</div>

<!-- Export Button -->
<form method="get" action="export_sales.php">
    <input type="hidden" name="start" value="<?= htmlspecialchars($start) ?>">
    <input type="hidden" name="end" value="<?= htmlspecialchars($end) ?>">
    <button type="submit">Export to CSV</button>
</form>

<form method="get" action="export_inventory_excel.php" style="margin-top:12px;">
    <button type="submit">Export Full Inventory Excel</button>
</form>

<h2>Sales in Range</h2>
<table>
    <tr>
        <th>Product</th>
        <th>Qty</th>
        <th>Total</th>
        <th>Date</th>
    </tr>
<?php
$totalQty = 0;
$totalIncome = 0;
while ($s = mysqli_fetch_assoc($sales)):
    $totalQty += (int)$s['quantity'];
    $totalIncome += (float)$s['total_price'];
?>
    <tr>
        <td><?= htmlspecialchars($s['name']) ?></td>
        <td><?= (int)$s['quantity'] ?></td>
        <td>₱<?= number_format($s['total_price'], 2) ?></td>
        <td><?= date('M d, Y', strtotime($s['date'])) ?></td>
    </tr>
<?php endwhile; ?>
    <tr>
        <th>Total</th>
        <th><?= $totalQty ?></th>
        <th>₱<?= number_format($totalIncome, 2) ?></th>
        <th></th>
    </tr>
</table>

<?php require_once 'footer.php'; ?>
