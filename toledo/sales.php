<?php
require_once 'header.php';
require_once 'excel_sync.php';

$message = '';
$products = mysqli_query($conn, "SELECT id, name, price, stock FROM products ORDER BY name");

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $product_id = (int)$_POST['product_id'];
    $qty = (int)$_POST['quantity'];
    $date = mysqli_real_escape_string($conn, $_POST['date']);
    $payment_method = mysqli_real_escape_string($conn, $_POST['payment_method']);

    $prod = mysqli_fetch_assoc(mysqli_query($conn, "SELECT price, stock FROM products WHERE id=$product_id"));
    if ($prod && $prod['stock'] >= $qty) {
        $total = $prod['price'] * $qty;
        mysqli_query($conn, "INSERT INTO sales (product_id, quantity, total_price, date) VALUES ($product_id,$qty,$total,'$date')");
        $sale_id = mysqli_insert_id($conn);
        mysqli_query($conn, "INSERT INTO payments (sale_id, payment_method, amount, date) VALUES ($sale_id,'$payment_method',$total,'$date')");
        mysqli_query($conn, "UPDATE products SET stock = stock - $qty WHERE id=$product_id");
        mysqli_query($conn, "INSERT INTO stock_logs (product_id, change_type, quantity, date) VALUES ($product_id,'sale',$qty,'$date')");
        $syncOk = syncInventoryWorkbook($conn);
        $message = $syncOk
            ? 'Sale recorded.'
            : 'Sale recorded. Excel sync failed. Close the workbook if it is open, then save again.';
    } else {
        $message = 'Not enough stock for this sale.';
    }
}

$recent = mysqli_query($conn, "SELECT s.id, p.name, s.quantity, s.total_price, s.date FROM sales s JOIN products p ON s.product_id=p.id ORDER BY s.date DESC, s.id DESC LIMIT 10");
?>
    <h1>Sales</h1>
    <?php if ($message): ?><div class="notice"><?php echo $message; ?></div><?php endif; ?>

    <form method="post">
        <label>Product</label>
        <select name="product_id">
            <?php while ($p = mysqli_fetch_assoc($products)): ?>
                <option value="<?php echo $p['id']; ?>">
                    <?php echo htmlspecialchars($p['name']); ?> (Stock: <?php echo $p['stock']; ?>)
                </option>
            <?php endwhile; ?>
        </select>
        <label>Quantity</label>
        <input type="number" name="quantity" required>
        <label>Date</label>
        <input type="date" name="date" value="<?php echo date('Y-m-d'); ?>" required>
        <label>Payment Method</label>
        <input type="text" name="payment_method" value="Cash" required>
        <button type="submit">Save Sale</button>
    </form>

    <h3>Recent Sales</h3>
    <table>
        <tr><th>ID</th><th>Product</th><th>Qty</th><th>Total</th><th>Date</th></tr>
        <?php while ($r = mysqli_fetch_assoc($recent)): ?>
            <tr>
                <td><?php echo $r['id']; ?></td>
                <td><?php echo htmlspecialchars($r['name']); ?></td>
                <td><?php echo (int)$r['quantity']; ?></td>
                <td>$<?php echo number_format($r['total_price'],2); ?></td>
                <td><?php echo $r['date']; ?></td>
            </tr>
        <?php endwhile; ?>
    </table>
<?php require_once 'footer.php'; ?>

