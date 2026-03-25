<?php
require_once 'header.php';
require_once 'excel_sync.php';

$message = $_GET['msg'] ?? '';

/*
|--------------------------------------------------------------------------
| ADD PRODUCT
|--------------------------------------------------------------------------
*/
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'add_product') {

    $name        = mysqli_real_escape_string($conn, $_POST['name']);
    $category_id = (int)$_POST['category_id'];
    $stock       = (int)$_POST['stock'];
    $labor_cost  = (float)$_POST['labor_cost'];

    $material_total = 0;
    $imagePath = '';

    if (isset($_FILES['image']) && $_FILES['image']['error'] === UPLOAD_ERR_OK) {
        $ext = pathinfo($_FILES['image']['name'], PATHINFO_EXTENSION);
        $newName = uniqid('equip_', true) . '.' . $ext;
        $uploadDir = 'uploads/';
        $imagePath = $uploadDir . $newName;
        move_uploaded_file($_FILES['image']['tmp_name'], $imagePath);
    }

    if (!empty($_POST['materials'])) {
        foreach ($_POST['materials'] as $material_id => $data) {
            if (!isset($data['use'])) continue;
            $qty = max(1, (int)$data['qty']);
            $res = mysqli_query(
                $conn,
                "SELECT price FROM materials WHERE material_id=".(int)$material_id
            );
            if ($row = mysqli_fetch_assoc($res)) {
                $material_total += $row['price'] * $qty;
            }
        }
    }

    $base_price  = $material_total + $labor_cost;
    $final_price = $base_price * 1.50;

    mysqli_query($conn, "
        INSERT INTO products
        (name, category_id, price, stock, labor_cost, material_cost, image)
        VALUES
        ('$name', $category_id, $final_price, $stock, $labor_cost, $material_total, '$imagePath')
    ");

    $product_id = mysqli_insert_id($conn);

    if (!empty($_POST['materials'])) {
        foreach ($_POST['materials'] as $material_id => $data) {
            if (!isset($data['use'])) continue;
            $qty = max(1, (int)$data['qty']);
            mysqli_query($conn, "
                INSERT INTO product_materials (product_id, material_id, qty)
                VALUES ($product_id, ".(int)$material_id.", $qty)
            ");
        }
    }

    $syncOk = syncInventoryWorkbook($conn);
    header('Location: products.php?msg=' . syncInventoryMessage('Product added', $syncOk));
    exit;
}

/*
|--------------------------------------------------------------------------
| EDIT PRODUCT
|--------------------------------------------------------------------------
*/
if ($_SERVER['REQUEST_METHOD'] === 'POST' && ($_POST['action'] ?? '') === 'edit_product') {

    $id          = (int)$_POST['id'];
    $name        = mysqli_real_escape_string($conn, $_POST['name']);
    $category_id = (int)$_POST['category_id'];
    $stock       = (int)$_POST['stock'];
    $labor_cost  = (float)$_POST['labor_cost'];

    $material_total = 0;

    mysqli_query($conn, "DELETE FROM product_materials WHERE product_id=$id");

    if (!empty($_POST['materials'])) {
        foreach ($_POST['materials'] as $material_id => $data) {

            if (!isset($data['use'])) continue;
            $qty = max(1, (int)$data['qty']);

            $res = mysqli_query(
                $conn,
                "SELECT price FROM materials WHERE material_id=".(int)$material_id
            );

            if ($row = mysqli_fetch_assoc($res)) {
                $material_total += $row['price'] * $qty;
            }

            mysqli_query($conn, "
                INSERT INTO product_materials (product_id, material_id, qty)
                VALUES ($id, ".(int)$material_id.", $qty)
            ");
        }
    }

    $base_price  = $material_total + $labor_cost;
    $final_price = $base_price * 1.50;

    mysqli_query($conn, "
        UPDATE products
        SET name='$name',
            category_id=$category_id,
            price=$final_price,
            stock=$stock,
            labor_cost=$labor_cost,
            material_cost=$material_total
        WHERE id=$id
    ");

    $syncOk = syncInventoryWorkbook($conn);
    header('Location: products.php?msg=' . syncInventoryMessage('Product updated', $syncOk));
    exit;
}

/*
|--------------------------------------------------------------------------
| DELETE PRODUCT
|--------------------------------------------------------------------------
*/
if (isset($_GET['delete'])) {
    $id = (int)$_GET['delete'];

    mysqli_query($conn, "DELETE FROM product_materials WHERE product_id=$id");
    mysqli_query($conn, "DELETE FROM products WHERE id=$id");

    $syncOk = syncInventoryWorkbook($conn);
    header('Location: products.php?msg=' . syncInventoryMessage('Product deleted', $syncOk));
    exit;
}

/*
|--------------------------------------------------------------------------
| FETCH DATA
|--------------------------------------------------------------------------
*/
$categories = mysqli_query($conn, "SELECT * FROM categories ORDER BY id");
$materials  = mysqli_query($conn, "SELECT * FROM materials ORDER BY material_name");

$products = mysqli_query($conn, "
    SELECT p.*, c.category_name
    FROM products p
    LEFT JOIN categories c ON c.id = p.category_id
    ORDER BY p.id DESC
");

/*
|--------------------------------------------------------------------------
| FETCH PRODUCT FOR EDIT
|--------------------------------------------------------------------------
*/
$editProduct = null;
$editMaterials = [];

if (isset($_GET['edit'])) {
    $editId = (int)$_GET['edit'];

    $editProduct = mysqli_fetch_assoc(
        mysqli_query($conn, "SELECT * FROM products WHERE id=$editId")
    );

    $pm = mysqli_query($conn, "
        SELECT material_id, qty
        FROM product_materials
        WHERE product_id=$editId
    ");

    while ($row = mysqli_fetch_assoc($pm)) {
        $editMaterials[$row['material_id']] = $row['qty'];
    }

    mysqli_data_seek($materials, 0);
    mysqli_data_seek($categories, 0);
}
?>

<h1><?= $editProduct ? 'Edit Equipment' : 'Add Equipment' ?></h1>
<?php if ($message): ?><div class="success"><?= htmlspecialchars($message) ?></div><?php endif; ?>

<form method="post" enctype="multipart/form-data">
    <input type="hidden" name="action"
           value="<?= $editProduct ? 'edit_product' : 'add_product' ?>">
    <?php if ($editProduct): ?>
        <input type="hidden" name="id" value="<?= $editProduct['id'] ?>">
    <?php endif; ?>

    <label>Name</label>
    <input type="text" name="name"
           value="<?= htmlspecialchars($editProduct['name'] ?? '') ?>" required>

    <label>Category</label>
    <select name="category_id" required>
        <?php while ($c = mysqli_fetch_assoc($categories)): ?>
            <option value="<?= $c['id'] ?>"
                <?= ($editProduct && $c['id']==$editProduct['category_id']) ? 'selected' : '' ?>>
                <?= htmlspecialchars($c['category_name']) ?>
            </option>
        <?php endwhile; ?>
    </select>

    <h4>Materials</h4>
    <div class="materials-box">
        <?php while ($m = mysqli_fetch_assoc($materials)):
            $used = isset($editMaterials[$m['material_id']]);
            $qty  = $used ? $editMaterials[$m['material_id']] : 1;
        ?>
        <div>
            <label>
                <input type="checkbox"
                       name="materials[<?= $m['material_id'] ?>][use]"
                       <?= $used ? 'checked' : '' ?>>
                <?= htmlspecialchars($m['material_name']) ?>
                (₱<?= number_format($m['price'],2) ?>)
            </label>
            Qty:
            <input type="number"
                   name="materials[<?= $m['material_id'] ?>][qty]"
                   value="<?= $qty ?>" min="1">
        </div>
        <?php endwhile; ?>
    </div>

        <label>Photo</label>
        <input type="file" name="image" accept="image/*">

        <label>Labor Cost</label>
        <input type="number" step="0.01" name="labor_cost"
            value="<?= $editProduct['labor_cost'] ?? '' ?>" required>

    <label>Stock</label>
    <input type="number" name="stock"
           value="<?= $editProduct['stock'] ?? '' ?>" required>

    <button type="submit">
        <?= $editProduct ? 'Update Product' : 'Save Product' ?>
    </button>

    <?php if ($editProduct): ?>
        <a href="products.php">Cancel</a>
    <?php endif; ?>
</form>

<hr>

<h3>Equipment List</h3>
<table>
<tr>
    <th>Image</th>
    <th>Name</th>
    <th>Category</th>
    <th>Price</th>
    <th>Stock</th>
    <th>Action</th>
</tr>
<?php while ($p = mysqli_fetch_assoc($products)): ?>
<tr>
    <td>
        <?php if (!empty($p['image'])): ?>
            <img src="<?= htmlspecialchars($p['image']) ?>" alt="Equipment Image" style="width:60px;height:60px;object-fit:cover;">
        <?php else: ?>
            <span>No Image</span>
        <?php endif; ?>
    </td>
    <td><?= htmlspecialchars($p['name']) ?></td>
    <td><?= htmlspecialchars($p['category_name']) ?></td>
    <td>₱<?= number_format($p['price'],2) ?></td>
    <td><?= $p['stock'] ?></td>
    <td>
        <a href="?edit=<?= $p['id'] ?>">Edit</a> |
        <a href="?delete=<?= $p['id'] ?>" onclick="return confirm('Delete?')">Delete</a>
    </td>
</tr>
<?php endwhile; ?>
</table>

<?php require_once 'footer.php'; ?>
