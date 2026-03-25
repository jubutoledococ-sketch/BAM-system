<?php
require_once 'header.php';
require_once 'excel_sync.php';

$message = isset($_GET['msg']) ? $_GET['msg'] : '';

// -----------------------------
// FETCH PRODUCTS & CUSTOMERS
// -----------------------------
$products = mysqli_query($conn, "
    SELECT id, name, stock, price
    FROM products
    WHERE stock > 0
    ORDER BY name
");

$customers = mysqli_query($conn, "
    SELECT id, customer_name
    FROM customers
    ORDER BY customer_name
");

// -----------------------------
// HANDLE POST REQUESTS
// -----------------------------
if ($_SERVER['REQUEST_METHOD'] === 'POST') {

    // -----------------------------
    // ADD PURCHASE (ADMIN)
    // -----------------------------
    if (($_POST['action'] ?? '') === 'add_purchase') {

        $product_id  = (int)$_POST['product_id'];
        $customer_id = ($_POST['customer_id'] !== '') ? (int)$_POST['customer_id'] : NULL;
        $quantity    = (int)$_POST['quantity'];
        $date        = mysqli_real_escape_string($conn, $_POST['date']);
        $status      = 'Pending';
        $discount    = floatval($_POST['discount'] ?? 0);

        $prod = mysqli_fetch_assoc(
            mysqli_query($conn, "SELECT price, stock FROM products WHERE id = $product_id")
        );

        if (!$prod || $quantity <= 0 || $quantity > $prod['stock']) {
            header("Location: purchases.php?msg=Invalid quantity.");
            exit;
        }

        $original_total = $prod['price'] * $quantity;
        $cost = $original_total - $discount;
        if ($cost < 0) $cost = 0;

        mysqli_query($conn, "
            INSERT INTO purchases
            (product_id, customer_id, quantity, original_cost, cost, date, status, created_by)
            VALUES (
                $product_id,
                " . ($customer_id ?? 'NULL') . ",
                $quantity,
                $original_total,
                $cost,
                '$date',
                '$status',
                'admin'
            )
        ");

        mysqli_query($conn, "
            UPDATE products
            SET stock = stock - $quantity
            WHERE id = $product_id
        ");

        mysqli_query($conn, "
            INSERT INTO stock_logs (product_id, change_type, quantity, date)
            VALUES ($product_id, 'sale', $quantity, '$date')
        ");

        $syncOk = syncInventoryWorkbook($conn);
        header("Location: purchases.php?msg=" . syncInventoryMessage('Purchase saved.', $syncOk));
        exit;
    }

    // -----------------------------
    // UPDATE STATUS
    // -----------------------------
    if (($_POST['action'] ?? '') === 'update_status') {
        $purchase_id = (int)$_POST['purchase_id'];
        $new_status  = mysqli_real_escape_string($conn, $_POST['status']);

        // Fetch current purchase BEFORE updating
        $current = mysqli_fetch_assoc(
            mysqli_query($conn, "SELECT * FROM purchases WHERE id = $purchase_id")
        );
        $old_status = $current['status'];

        // Update status
        $update = mysqli_query($conn, "
            UPDATE purchases
            SET status = '$new_status'
            WHERE id = $purchase_id
        ");
        if (!$update) {
            die("Update failed: " . mysqli_error($conn));
        }

        // Trigger Delivered logic once
        if ($new_status === 'Delivered' && $old_status !== 'Delivered') {
            mysqli_query($conn, "
                INSERT INTO sales (purchase_id, product_id, quantity, total_price, date)
                VALUES (
                    {$current['id']},
                    {$current['product_id']},
                    {$current['quantity']},
                    {$current['cost']},
                    '{$current['date']}'
                )
            ");

            $sale_id = mysqli_insert_id($conn);

            mysqli_query($conn, "
                INSERT INTO payments (sale_id, payment_method, amount, date)
                VALUES (
                    $sale_id,
                    'Cash',
                    {$current['cost']},
                    '{$current['date']}'
                )
            ");
        }

        $syncOk = syncInventoryWorkbook($conn);
        header("Location: purchases.php?msg=" . syncInventoryMessage('Status updated.', $syncOk));
        exit;
    }

    // -----------------------------
    // EDIT TOTAL (DISCOUNT)
    // -----------------------------
    if (($_POST['action'] ?? '') === 'edit_total') {
        $purchase_id = (int)$_POST['purchase_id'];
        $new_total   = (float)$_POST['new_total'];

        mysqli_query($conn, "
            UPDATE purchases
            SET cost = $new_total
            WHERE id = $purchase_id
        ");

        $syncOk = syncInventoryWorkbook($conn);
        header("Location: purchases.php?msg=" . syncInventoryMessage('Total updated.', $syncOk));
        exit;
    }
}

// -----------------------------
// FETCH RECENT PURCHASES (exclude Delivered/Repair)
// -----------------------------
$recent = mysqli_query($conn, "
    SELECT
        pu.id,
        pr.name AS product_name,
        c.customer_name,
        pu.quantity,
        pu.cost,
        pu.original_cost,
        pu.date,
        pu.status
    FROM purchases pu
    JOIN products pr ON pu.product_id = pr.id
    LEFT JOIN customers c ON pu.customer_id = c.id
    WHERE pu.status NOT IN ('Delivered','Repair')
    ORDER BY pu.id DESC
    LIMIT 20
");
?>

<h1>Order Equipment</h1>
<form method="post" action="purchases.php">
    <input type="hidden" name="action" value="add_purchase">

    <label>Equipment</label>
    <select name="product_id" id="product-select" required>
    <option value="">Select Equipment</option>

    <?php while ($p = mysqli_fetch_assoc($products)): ?>
        <option value="<?php echo $p['id']; ?>"
                data-stock="<?php echo $p['stock']; ?>"
                data-price="<?php echo $p['price']; ?>">
            <?php echo htmlspecialchars($p['name']); ?>
        </option>
    <?php endwhile; ?>
    </select>


    <label>Price (₱)</label>
        <input type="text" id="price-display" readonly>


    <label>Quantity</label>
    <select name="quantity" id="quantity-select" required>
        <option value="">Select Quantity</option>
    </select>

    <label>Client</label>
    <select name="customer_id">
        <option value="">Walk-in / Select Client</option>
        <?php mysqli_data_seek($customers, 0); ?>
        <?php while ($c = mysqli_fetch_assoc($customers)): ?>
            <option value="<?php echo $c['id']; ?>">
                <?php echo htmlspecialchars($c['customer_name']); ?>
            </option>
        <?php endwhile; ?>
    </select>

    <label>Date</label>
    <input type="date" name="date" value="<?php echo date('Y-m-d'); ?>" required>

    <label>Payment Method</label>
    <select name="payment_method" required>
        <option value="Cash">Cash</option>
        <option value="Gcash">Gcash</option>
        <option value="Bank Transfer">Bank Transfer</option>
    </select>

    <label>Discount (₱)</label>
    <input type="number" step="0.01" name="discount" value="0">

    <label>Total (₱)</label>
<input type="text" id="total-display" readonly>


    <button type="submit">Save Purchase</button>
</form>


<!-- RECENT ORDERS -->
<h3>Recent Orders</h3>
<table>
<tr>
    <th>Client</th>
    <th>Equipment</th>
    <th>Qty</th>
    <th>Total</th>
    <th>Date</th>
    <th>Status</th>
    <th>Discount</th>
</tr>

<?php while ($r = mysqli_fetch_assoc($recent)): ?>
<tr>
    <td><?php echo $r['customer_name'] ?: 'Walk-in'; ?></td>
    <td><?php echo htmlspecialchars($r['product_name']); ?></td>
    <td><?php echo (int)$r['quantity']; ?></td>
    <td>
        ₱<?php echo number_format($r['cost'],2); ?>
        <?php if ($r['cost'] < $r['original_cost']): ?>
            <span style="color:red;font-size:12px;">(discounted)</span>
        <?php endif; ?>
    </td>
    <td><?php echo $r['date']; ?></td>
    <td>
        <form method="post">
            <input type="hidden" name="action" value="update_status">
            <input type="hidden" name="purchase_id" value="<?php echo $r['id']; ?>">
            <select name="status" onchange="this.form.submit()">
                <?php foreach (['Pending','Fabricating','Quality Check','Delivered','Repair'] as $s): ?>
                    <option value="<?php echo $s; ?>" <?php if ($r['status']===$s) echo 'selected'; ?>>
                        <?php echo $s; ?>
                    </option>
                <?php endforeach; ?>
            </select>
        </form>
    </td>
    <td><?php echo number_format($r['original_cost'] - $r['cost'],2); ?></td>
</tr>
<?php endwhile; ?>
</table>

<script>
const productSelect  = document.getElementById('product-select');
const quantitySelect = document.getElementById('quantity-select');
const priceDisplay   = document.getElementById('price-display');
const totalDisplay   = document.getElementById('total-display');
const discountInput  = document.querySelector('input[name="discount"]');

let selectedPrice = 0;

function updateTotal() {
    const qty = parseInt(quantitySelect.value || 0);
    const discount = parseFloat(discountInput.value || 0);
    let total = (selectedPrice * qty) - discount;
    if (total < 0) total = 0;
    totalDisplay.value = total.toFixed(2);
}

productSelect.addEventListener('change', function () {
    const opt = this.selectedOptions[0];
    const stock = parseInt(opt?.dataset.stock || 0);
    selectedPrice = parseFloat(opt?.dataset.price || 0);

    priceDisplay.value = selectedPrice.toFixed(2);

    quantitySelect.innerHTML = '<option value="">Select Quantity</option>';
    for (let i = 1; i <= stock; i++) {
        quantitySelect.innerHTML += `<option value="${i}">${i}</option>`;
    }

    updateTotal();
});

quantitySelect.addEventListener('change', updateTotal);
discountInput.addEventListener('input', updateTotal);
</script>


<?php require_once 'footer.php'; ?>
