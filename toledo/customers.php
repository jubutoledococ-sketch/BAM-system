<?php
require_once 'header.php';
require_once 'excel_sync.php';

$message = $_GET['msg'] ?? '';

// Fetch all customers
$customers = mysqli_query($conn, "SELECT * FROM customers ORDER BY id DESC");

// Check if editing
$editCustomer = null;
if (isset($_GET['edit'])) {
    $editId = (int)$_GET['edit'];
    $editCustomer = mysqli_fetch_assoc(
        mysqli_query($conn, "SELECT * FROM customers WHERE id=$editId")
    );
}

// Handle POST requests
if ($_SERVER['REQUEST_METHOD'] === 'POST') {

    $action = $_POST['action'] ?? '';

    $customer_name = mysqli_real_escape_string($conn, $_POST['customer_name']);
    $contact       = mysqli_real_escape_string($conn, $_POST['contact']);
    $email         = mysqli_real_escape_string($conn, $_POST['email']);
    $address       = mysqli_real_escape_string($conn, $_POST['address']);
    $client_type   = mysqli_real_escape_string($conn, $_POST['client_type']);
    $payment_terms = mysqli_real_escape_string($conn, $_POST['payment_terms']);

    if ($action === 'add_customer') {
        mysqli_query($conn, "
            INSERT INTO customers
            (customer_name, contact, email, address, client_type, payment_terms, created_at)
            VALUES
            ('$customer_name', '$contact', '$email', '$address', '$client_type', '$payment_terms', CURRENT_DATE)
        ");
        $syncOk = syncInventoryWorkbook($conn);
        header('Location: customers.php?msg=' . syncInventoryMessage('Customer added.', $syncOk));
        exit;
    }

    if ($action === 'edit_customer') {
        $id = (int)$_POST['id'];
        mysqli_query($conn, "
            UPDATE customers SET
            customer_name='$customer_name',
            contact='$contact',
            email='$email',
            address='$address',
            client_type='$client_type',
            payment_terms='$payment_terms'
            WHERE id=$id
        ");
        $syncOk = syncInventoryWorkbook($conn);
        header('Location: customers.php?msg=' . syncInventoryMessage('Customer updated.', $syncOk));
        exit;
    }
}

// Handle delete
if (isset($_GET['delete'])) {
    $id = (int)$_GET['delete'];
    mysqli_query($conn, "DELETE FROM customers WHERE id=$id");
    $syncOk = syncInventoryWorkbook($conn);
    header('Location: customers.php?msg=' . syncInventoryMessage('Customer deleted.', $syncOk));
    exit;
}
?>

<h1>Client Information</h1>
<?php if ($message): ?>
    <div class="success"><?= htmlspecialchars($message) ?></div>
<?php endif; ?>

<!-- SINGLE FORM FOR ADD / EDIT -->
<form method="post">
    <input type="hidden" name="action" value="<?= $editCustomer ? 'edit_customer' : 'add_customer' ?>">
    <?php if ($editCustomer): ?>
        <input type="hidden" name="id" value="<?= $editCustomer['id'] ?>">
    <?php endif; ?>

    <label>Customer Name</label>
    <input type="text" name="customer_name" required
           value="<?= htmlspecialchars($editCustomer['customer_name'] ?? '') ?>">

    <label>Contact Number</label>
    <input type="text" name="contact"
           value="<?= htmlspecialchars($editCustomer['contact'] ?? '') ?>">

    <label>Email</label>
    <input type="email" name="email"
           value="<?= htmlspecialchars($editCustomer['email'] ?? '') ?>">

    <label>Address</label>
    <input type="text" name="address"
           value="<?= htmlspecialchars($editCustomer['address'] ?? '') ?>">

    <label>Client Type</label>
    <select name="client_type">
        <option value="Walk-in" <?= ($editCustomer['client_type'] ?? '') === 'Walk-in' ? 'selected' : '' ?>>Walk-in</option>
        <option value="PhoneCall" <?= ($editCustomer['client_type'] ?? '') === 'PhoneCall' ? 'selected' : '' ?>>Phone Call</option>
    </select>

    <label>Payment Terms</label>
    <select name="payment_terms">
        <option value="COH" <?= ($editCustomer['payment_terms'] ?? '') === 'COH' ? 'selected' : '' ?>>Full Payment</option>
        <option value="Downpayment" <?= ($editCustomer['payment_terms'] ?? '') === 'Downpayment' ? 'selected' : '' ?>>Down Payment</option>
    </select>

    <button type="submit"><?= $editCustomer ? 'Update Customer' : 'Save Customer' ?></button>
    <?php if ($editCustomer): ?>
        <a href="customers.php">Cancel</a>
    <?php endif; ?>
</form>

<hr>

<h3>Client List</h3>
<table>
    <tr>
        <th>Client Name</th>
        <th>Contact</th>
        <th>Email</th>
        <th>Address</th>
        <th>Client Type</th>
        <th>Payment Terms</th>
        <th>Date Added</th>
        <th>Actions</th>
    </tr>

    <?php while ($c = mysqli_fetch_assoc($customers)): ?>
        <tr>
            <td><?= htmlspecialchars($c['customer_name']) ?></td>
            <td><?= htmlspecialchars($c['contact']) ?></td>
            <td><?= htmlspecialchars($c['email']) ?></td>
            <td><?= htmlspecialchars($c['address']) ?></td>
            <td><?= htmlspecialchars($c['client_type']) ?></td>
            <td><?= htmlspecialchars($c['payment_terms']) ?></td>
            <td><?= date('M d, Y', strtotime($c['created_at'])) ?></td>
            <td>
                <a href="customers.php?edit=<?= $c['id'] ?>">Edit</a> |
                <a href="customers.php?delete=<?= $c['id'] ?>"
                   onclick="return confirm('Delete customer?');">Delete</a>
            </td>
        </tr>
    <?php endwhile; ?>
</table>

<?php require_once 'footer.php'; ?>
