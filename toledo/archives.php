<?php
require_once 'header.php';

$message = isset($_GET['msg']) ? $_GET['msg'] : '';

// Fetch archived purchases (Delivered & Repair) sorted by client name A-Z
$archives = mysqli_query($conn, "
    SELECT pu.id, pr.name AS product_name, c.customer_name, pu.quantity, pu.cost, pu.date, pu.status
    FROM purchases pu
    JOIN products pr ON pu.product_id = pr.id
    LEFT JOIN customers c ON pu.customer_id = c.id
    WHERE pu.status IN ('Delivered', 'Repair')
    ORDER BY c.customer_name ASC, pu.id ASC
");
?>

<h1>Archived Purchases</h1>

<?php if ($message): ?>
    <div class="success"><?php echo htmlspecialchars($message); ?></div>
<?php endif; ?>

<table>
<tr>
    <th>Client</th>
    <th>Equipment</th>
    <th>Quantity</th>
    <th>Total</th>
    <th>Date</th>
    <th>Status</th>
</tr>

<?php while ($r = mysqli_fetch_assoc($archives)): ?>
<tr>
    <td><?php echo $r['customer_name'] ?: 'Walk-in'; ?></td>
    <td><?php echo htmlspecialchars($r['product_name']); ?></td>
    <td><?php echo (int)$r['quantity']; ?></td>
    <td>₱<?php echo number_format($r['cost'],2); ?></td>
    <td><?php echo $r['date']; ?></td>
    <td><?php echo $r['status']; ?></td>
</tr>
<?php endwhile; ?>
</table>

<?php require_once 'footer.php'; ?>
