<?php
require_once 'sidebar.php';

// Fetch counts for interactive cards
$newJobsCount = mysqli_fetch_assoc(mysqli_query($conn, "SELECT COUNT(*) AS cnt FROM purchases WHERE status='Pending'"))['cnt'];
$inProgressCount = mysqli_fetch_assoc(mysqli_query($conn, "SELECT COUNT(*) AS cnt FROM purchases WHERE status='Fabricating'"))['cnt'];

$dailySales = mysqli_fetch_assoc(mysqli_query($conn, "SELECT IFNULL(SUM(total_price),0) AS amt FROM sales WHERE date = CURDATE()"));
$lowStock = mysqli_query($conn, "SELECT name, stock FROM products WHERE stock < 5 ORDER BY stock ASC");
$recentSales = mysqli_query($conn, "SELECT id, total_price, date 
                                   FROM sales 
                                   ORDER BY date DESC 
                                   LIMIT 5");

// Recent fabrication updates (last 5)
$recentFabrications = mysqli_query($conn, "
    SELECT pu.id, pr.name AS product_name, c.customer_name, pu.status, pu.date
    FROM purchases pu
    LEFT JOIN products pr ON pu.product_id = pr.id
    LEFT JOIN customers c ON pu.customer_id = c.id
    ORDER BY pu.date DESC, pu.id DESC
    LIMIT 5
");
?>

<h1>Welcome Staff!</h1>
<div class="grid" style="margin-bottom:20px; gap:15px; display:flex; flex-wrap:wrap;">
    <!-- New Job -->
    <div class="card" style="flex:1; min-width:200px; background-color:#b1a07d; color:white; padding:20px; border-radius:8px;">
        <h3>New Job</h3>
        <p><?php echo $newJobsCount; ?> pending jobs</p>
    </div>

    <!-- Recent Fabrication Updates -->
    <div class="card" style="flex:1; min-width:200px; background-color:#8c8c8c; color:white; padding:20px; border-radius:8px;">
        <h3>Recent Updates</h3>
        <p>Last 5 fabrication orders</p>
    </div>

    <!-- In-Progress Orders -->
    <div class="card" style="flex:1; min-width:200px; background-color:#7f8c8d; color:white; padding:20px; border-radius:8px;">
        <h3>In-Progress Orders</h3>
        <p><?php echo $inProgressCount; ?> orders being fabricated</p>
    </div>
</div>


<!-- Daily Sales -->
<div class="card" style="background-color:#666666;color:white; margin-bottom:20px;">
    <h3>Daily Sales</h3>
    <p>₱<?php echo number_format($dailySales['amt'], 2); ?></p>
</div>

<!-- Low Stock Alerts -->
<h2>Low Stock Products</h2>
<table>
    <tr>
        <th>Product</th>
        <th>Stock</th>
    </tr>
    <?php while($product = mysqli_fetch_assoc($lowStock)): ?>
        <tr>
            <td><?php echo $product['name']; ?></td>
            <td><?php echo $product['stock']; ?></td>
        </tr>
    <?php endwhile; ?>
</table>

<!-- Recent Sales -->
<h2>Recent Sales</h2>
<table>
    <tr>
        <th>Total Price</th>
        <th>Date</th>
    </tr>
    <?php while($sale = mysqli_fetch_assoc($recentSales)): ?>
        <tr>
            <td>₱<?php echo number_format($sale['total_price'], 2); ?></td>
            <td><?php echo $sale['date']; ?></td>
        </tr>
    <?php endwhile; ?>
</table>

<?php require_once 'footer.php'; ?>
