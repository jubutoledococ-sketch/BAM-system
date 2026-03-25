<?php
require_once 'header.php';
require_once 'excel_sync.php';

$message = $_GET['msg'] ?? '';

// Handle POST requests
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $action = $_POST['action'] ?? '';
    $category_name = mysqli_real_escape_string($conn, $_POST['category_name']);

    if ($action === 'add_category' && $category_name !== '') {
        mysqli_query($conn, "INSERT INTO categories (category_name) VALUES ('$category_name')");
        $syncOk = syncInventoryWorkbook($conn);
        header('Location: categories.php?msg=' . syncInventoryMessage('Category added.', $syncOk));
        exit;
    }

    if ($action === 'edit_category') {
        $id = (int)$_POST['id'];
        mysqli_query($conn, "UPDATE categories SET category_name='$category_name' WHERE id=$id");
        $syncOk = syncInventoryWorkbook($conn);
        header('Location: categories.php?msg=' . syncInventoryMessage('Category updated.', $syncOk));
        exit;
    }
}

// Handle delete
if (isset($_GET['delete'])) {
    $id = (int)$_GET['delete'];
    mysqli_query($conn, "DELETE FROM categories WHERE id=$id");
    $syncOk = syncInventoryWorkbook($conn);
    header('Location: categories.php?msg=' . syncInventoryMessage('Category deleted.', $syncOk));
    exit;
}

// Fetch categories
$categories = mysqli_query($conn, "SELECT * FROM categories ORDER BY id ASC");

// Check if editing
$editCategory = null;
if (isset($_GET['edit'])) {
    $editId = (int)$_GET['edit'];
    $editCategory = mysqli_fetch_assoc(
        mysqli_query($conn, "SELECT * FROM categories WHERE id=$editId")
    );
}
?>

<h1>Equipment Categories</h1>

<?php if ($message): ?>
    <div class="success"><?= htmlspecialchars($message) ?></div>
<?php endif; ?>

<!-- SINGLE FORM FOR ADD / EDIT -->
<form method="post">
    <input type="hidden" name="action" value="<?= $editCategory ? 'edit_category' : 'add_category' ?>">
    <?php if ($editCategory): ?>
        <input type="hidden" name="id" value="<?= $editCategory['id'] ?>">
    <?php endif; ?>

    <label>Category Name</label>
    <input type="text" name="category_name" required
           value="<?= htmlspecialchars($editCategory['category_name'] ?? '') ?>">

    <button type="submit"><?= $editCategory ? 'Update Category' : 'Save Category' ?></button>
    <?php if ($editCategory): ?>
        <a href="categories.php">Cancel</a>
    <?php endif; ?>
</form>

<hr>

<h3>Category List</h3>
<table>
    <tr>
        <th>Name</th>
        <th>Actions</th>
    </tr>
    <?php while ($c = mysqli_fetch_assoc($categories)): ?>
        <tr>
            <td><?= htmlspecialchars($c['category_name']) ?></td>
            <td>
                <a href="categories.php?edit=<?= $c['id'] ?>">Edit</a> |
                <a href="categories.php?delete=<?= $c['id'] ?>"
                   onclick="return confirm('Delete category?');">Delete</a>
            </td>
        </tr>
    <?php endwhile; ?>
</table>

<?php require_once 'footer.php'; ?>
