<?php
require_once 'sidebar.php';

// -----------------------------
// FLASH MESSAGE
// -----------------------------
$message = $_SESSION['msg'] ?? '';
unset($_SESSION['msg']);

// -----------------------------
// HANDLE FORM SUBMIT
// -----------------------------
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'add_purchase') {

    $product_id     = (int) $_POST['product_id'];
    $customer_id    = ($_POST['customer_id'] !== '') ? (int) $_POST['customer_id'] : null;
    $qty            = (int) $_POST['quantity'];
    $discount       = (float) ($_POST['discount'] ?? 0);
    $date           = mysqli_real_escape_string($conn, $_POST['date']);
    $payment_method = mysqli_real_escape_string($conn, $_POST['payment_method']);

    $prod = mysqli_fetch_assoc(
        mysqli_query($conn, "SELECT price, stock FROM products WHERE id = $product_id")
    );

    if ($prod && $qty > 0 && $prod['stock'] >= $qty) {

        $original_cost = $prod['price'] * $qty;
        $final_cost    = max(0, $original_cost - $discount);

        mysqli_query($conn, "
            INSERT INTO purchases
            (product_id, customer_id, quantity, cost, original_cost, date, status, created_by)
            VALUES
            ($product_id, " . ($customer_id ?? 'NULL') . ", $qty, $final_cost, $original_cost, '$date', 'Pending', 'staff')
        ");

        mysqli_query($conn, "UPDATE products SET stock = stock - $qty WHERE id = $product_id");
        mysqli_query($conn, "INSERT INTO stock_logs (product_id, change_type, quantity, date)
                             VALUES ($product_id, 'sale', $qty, '$date')");

        mysqli_query($conn, "
            INSERT INTO sales (product_id, quantity, total_price, date)
            VALUES ($product_id, $qty, $final_cost, '$date')
        ");

        $sale_id = mysqli_insert_id($conn);

        mysqli_query($conn, "
            INSERT INTO payments (sale_id, payment_method, amount, date)
            VALUES ($sale_id, '$payment_method', $final_cost, '$date')
        ");

        $_SESSION['msg'] = "Purchase saved successfully.";
    } else {
        $_SESSION['msg'] = "Not enough stock or invalid quantity.";
    }

    // 🔥 THIS IS THE IMPORTANT PART
    header('Location: staff_purchases.php');
    exit;
}

// -----------------------------
// FETCH DATA (AFTER POST)
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

$recent = mysqli_query($conn, "
    SELECT 
        p.quantity, p.cost, p.original_cost, p.date, p.status,
        pr.name AS product_name,
        c.customer_name
    FROM purchases p
    JOIN products pr ON p.product_id = pr.id
    LEFT JOIN customers c ON p.customer_id = c.id
    WHERE p.status NOT IN ('Delivered', 'Repair')
    ORDER BY p.id DESC
    LIMIT 20
");
?>

<h1>Order Equipment</h1>

<?php if ($message): ?>
    <p class="alert"><?php echo htmlspecialchars($message); ?></p>
<?php endif; ?>

<form method="post">
    <input type="hidden" name="action" value="add_purchase">

    <label>Equipment</label>
    <select name="product_id" id="product-select" required>
        <option value="">Select Equipment</option>
        <?php while ($p = mysqli_fetch_assoc($products)): ?>
            <option value="<?= $p['id'] ?>"
                    data-stock="<?= $p['stock'] ?>"
                    data-price="<?= $p['price'] ?>">
                <?= htmlspecialchars($p['name']) ?>
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
        <?php while ($c = mysqli_fetch_assoc($customers)): ?>
            <option value="<?= $c['id'] ?>">
                <?= htmlspecialchars($c['customer_name']) ?>
            </option>
        <?php endwhile; ?>
    </select>

    <label>Date</label>
    <input type="date" name="date" value="<?= date('Y-m-d') ?>" required>

    <label>Payment Method</label>
    <select name="payment_method" required>
        <option>Cash</option>
        <option>Gcash</option>
        <option>Bank Transfer</option>
    </select>

    <label>Discount (₱)</label>
    <input type="number" step="0.01" name="discount" value="0">

    <label>Total (₱)</label>
    <input type="text" id="total-display" readonly>

    <button type="submit">Save Purchase</button>
</form>

<h3>Recent Orders</h3>

<table>
<tr>
    <th>Client</th>
    <th>Equipment</th>
    <th>Qty</th>
    <th>Total</th>
    <th>Date</th>
    <th>Status</th>
</tr>

<?php while ($r = mysqli_fetch_assoc($recent)): ?>
<tr>
    <td><?= $r['customer_name'] ?: 'Walk-in' ?></td>
    <td><?= htmlspecialchars($r['product_name']) ?></td>
    <td><?= (int)$r['quantity'] ?></td>
    <td>
        ₱<?= number_format($r['cost'], 2) ?>
        <?php if ($r['cost'] < $r['original_cost']): ?>
            <span style="color:red;font-size:12px;">(discounted)</span>
        <?php endif; ?>
    </td>
    <td><?= $r['date'] ?></td>
    <td><?= htmlspecialchars($r['status']) ?></td>
</tr>
<?php endwhile; ?>
</table>

<script>
const productSelect  = document.getElementById('product-select');
const quantitySelect = document.getElementById('quantity-select');
const priceDisplay   = document.getElementById('price-display');
const totalDisplay   = document.getElementById('total-display');
const discountInput  = document.querySelector('[name="discount"]');

let selectedPrice = 0;

function updateTotal() {
    const qty = parseInt(quantitySelect.value || 0);
    const discount = parseFloat(discountInput.value || 0);
    totalDisplay.value = Math.max(0, (selectedPrice * qty) - discount).toFixed(2);
}

productSelect.addEventListener('change', () => {
    const opt = productSelect.selectedOptions[0];
    selectedPrice = parseFloat(opt.dataset.price || 0);
    priceDisplay.value = selectedPrice.toFixed(2);

    quantitySelect.innerHTML = '<option value="">Select Quantity</option>';
    for (let i = 1; i <= opt.dataset.stock; i++) {
        quantitySelect.innerHTML += `<option value="${i}">${i}</option>`;
    }
    updateTotal();
});

quantitySelect.addEventListener('change', updateTotal);
discountInput.addEventListener('input', updateTotal);
</script>

<?php require_once 'footer.php'; ?>
