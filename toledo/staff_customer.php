
<?php 
require_once 'sidebar.php'; 

// Flash message
$message = isset($_GET['msg']) ? $_GET['msg'] : '';
$search  = isset($_GET['search']) ? trim($_GET['search']) : '';

// -----------------------------
// HANDLE ADD CUSTOMER
// -----------------------------
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'add_customer') {
    $customer_name = mysqli_real_escape_string($conn, $_POST['customer_name']);
    $contact       = mysqli_real_escape_string($conn, $_POST['contact']);
    $email         = mysqli_real_escape_string($conn, $_POST['email']);
    $address       = mysqli_real_escape_string($conn, $_POST['address']);
    $client_type   = mysqli_real_escape_string($conn, $_POST['client_type']);
    $payment_terms = mysqli_real_escape_string($conn, $_POST['payment_terms']);

    mysqli_query($conn, "INSERT INTO customers 
        (customer_name, contact, email, address, client_type, payment_terms, created_at) 
        VALUES ('$customer_name', '$contact', '$email', '$address', '$client_type', '$payment_terms', NOW())"
    );

    header('Location: staff_customer.php?msg=Customer+added.');
    exit;
}

// -----------------------------
// FETCH CUSTOMERS (with optional search)
// -----------------------------
$sql = "SELECT * FROM customers WHERE 1";
if ($search !== '') {
    $search_escaped = mysqli_real_escape_string($conn, $search);
    $sql .= " AND (customer_name LIKE '%$search_escaped%' 
                 OR contact LIKE '%$search_escaped%' 
                 OR email LIKE '%$search_escaped%')";
}
$sql .= " ORDER BY id DESC";

$customers = mysqli_query($conn, $sql);
?>

<h1>Clients</h1>

<?php if ($message): ?>
    <div class="success"><?php echo htmlspecialchars($message); ?></div>
<?php endif; ?>

<!-- -----------------------------
     ADD NEW CUSTOMER FORM
----------------------------- -->
<h3>Add New Customer</h3>
<form method="post">
    <input type="hidden" name="action" value="add_customer">

    <label>Customer Name</label>
    <input type="text" name="customer_name" required>

    <label>Contact Number</label>
    <input type="text" name="contact">

    <label>Email</label>
    <input type="email" name="email">

    <label>Address</label>
    <input type="text" name="address">

    <label>Client Type</label>
    <select name="client_type">
        <option value="Walk-in">Walk-in</option>
        <option value="PhoneCall">Phone Call</option>
    </select>

    <label>Payment Terms</label>
    <select name="payment_terms">
        <option value="COH">Full Payment</option>
        <option value="Downpayment">Down Payment</option>
    </select>

    <button type="submit">Save Customer</button>
</form>


<h1>Client List</h1>


<h3>Search Clients</h3>
<form method="get">
    <input type="text" name="search"
           placeholder="Search by name, contact, or email"
           value="<?php echo htmlspecialchars($search); ?>">
    <button>Search</button>
</form>
<table>
    <tr>
        <th>Client Name</th>
        <th>Contact</th>
        <th>Email</th>
        <th>Address</th>
        <th>Client Type</th>
        <th>Payment Terms</th>
        <th>Date Added</th>
    </tr>
    <?php if (mysqli_num_rows($customers) === 0): ?>
        <tr>
            <td colspan="7" style="text-align:center;">No clients found.</td>
        </tr>
    <?php else: ?>
        <?php while ($c = mysqli_fetch_assoc($customers)): ?>
            <tr>
                <td><?php echo htmlspecialchars($c['customer_name']); ?></td>
                <td><?php echo htmlspecialchars($c['contact']); ?></td>
                <td><?php echo htmlspecialchars($c['email']); ?></td>
                <td><?php echo htmlspecialchars($c['address']); ?></td>
                <td><?php echo htmlspecialchars($c['client_type']); ?></td>
                <td><?php echo htmlspecialchars($c['payment_terms']); ?></td>
                <td><?php echo date('M d, Y', strtotime($c['created_at'])); ?></td>
            </tr>
        <?php endwhile; ?>
    <?php endif; ?>
</table>

<?php require_once 'footer.php'; ?>
