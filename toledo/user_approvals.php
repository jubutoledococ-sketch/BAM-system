<?php
require_once 'header.php';
require_once 'db.php';

// Handle user approval
$approveMsg = '';
if (isset($_POST['approve_user'])) {
    $approve_id = intval($_POST['approve_id']);
    $stmt = $conn->prepare("UPDATE users SET approved = 1 WHERE id = ?");
    $stmt->bind_param('i', $approve_id);
    if ($stmt->execute()) {
        $approveMsg = '<span style="color:green;">User approved successfully!</span>';
    } else {
        $approveMsg = '<span style="color:red;">Error approving user.</span>';
    }
    $stmt->close();
}

// Fetch unapproved users
$pendingUsers = $conn->query("SELECT id, username, name, age, address, role FROM users WHERE approved = 0");
?>

<h1>Pending User Approvals</h1>
<?php if (!empty($approveMsg)) echo $approveMsg; ?>
<?php if ($pendingUsers && $pendingUsers->num_rows > 0): ?>
    <table style="width:100%;border-collapse:collapse;">
        <tr style="background:#f6f6f6;">
            <th>Name</th>
            <th>Username</th>
            <th>Age</th>
            <th>Address</th>
            <th>Role</th>
            <th>Action</th>
        </tr>
        <?php while($user = $pendingUsers->fetch_assoc()): ?>
        <tr>
            <td><?php echo htmlspecialchars($user['name']); ?></td>
            <td><?php echo htmlspecialchars($user['username']); ?></td>
            <td><?php echo htmlspecialchars($user['age']); ?></td>
            <td><?php echo htmlspecialchars($user['address']); ?></td>
            <td><?php echo htmlspecialchars($user['role']); ?></td>
            <td>
                <form method="post" style="margin:0;">
                    <input type="hidden" name="approve_id" value="<?php echo $user['id']; ?>">
                    <button type="submit" name="approve_user" style="background:#52c41a;color:#fff;padding:5px 10px;border:none;border-radius:3px;">Approve</button>
                </form>
            </td>
        </tr>
        <?php endwhile; ?>
    </table>
<?php else: ?>
    <p>No users pending approval.</p>
<?php endif; ?>

<?php require_once 'footer.php'; ?>
