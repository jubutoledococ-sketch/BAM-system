<?php
require_once 'sidebar.php';

$search = trim($_GET['search'] ?? '');

$sql = "
SELECT p.*, c.category_name
FROM products p
LEFT JOIN categories c ON p.category_id = c.id
";

if ($search !== '') {
    $safe = mysqli_real_escape_string($conn, $search);
    $sql .= " WHERE p.name LIKE '%$safe%'
              OR c.category_name LIKE '%$safe%'";
}

$sql .= " ORDER BY p.id DESC";
$products = mysqli_query($conn, $sql);
?>

<h1>Equipment List</h1>

<form method="get">
    <input type="text" name="search"
           placeholder="Search equipment or category"
           value="<?php echo htmlspecialchars($search); ?>">
    <button>Search</button>
</form>

<table>
<tr>
    <th>Image</th>
    <th>ID</th>
    <th>Name</th>
    <th>Category</th>
    <th>Price</th>
    <th>Stock</th>
</tr>

<?php while ($p = mysqli_fetch_assoc($products)): ?>
<tr>
    <td><img src="12.jpg" alt="Equipment Image" style="width:60px;height:60px;object-fit:cover;"></td>
    <td><?php echo $p['id']; ?></td>
    <td><?php echo htmlspecialchars($p['name']); ?></td>
    <td><?php echo htmlspecialchars($p['category_name']); ?></td>
    <td>
        ₱<?php echo number_format($p['price'], 2); ?>
        <?php if ((float)$p['original_price'] > (float)$p['price']): ?>
        <span style="color:red;font-weight:bold;">(Discounted)</span>
        <?php endif; ?>

    </td>
    <td><?php echo (int)$p['stock']; ?></td>
</tr>
<?php endwhile; ?>
</table>

<?php require_once 'footer.php'; ?>
