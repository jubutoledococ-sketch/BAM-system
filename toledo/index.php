<?php
session_start();
require_once "db.php";

$error = "";
$registerMsg = "";

if (isset($_POST['login'])) {

    $username = trim($_POST['username']);
    $password = trim($_POST['password']);

    $stmt = $conn->prepare(
        "SELECT username, password, role FROM users WHERE username = ?"
    );
    $stmt->bind_param("s", $username);
    $stmt->execute();
    $result = $stmt->get_result();

    if ($result && $result->num_rows === 1) {
        $user = $result->fetch_assoc();

        // Plain text password comparison (not secure)
        if ($password === $user['password']) {

            $_SESSION['username'] = $user['username'];
            $_SESSION['role'] = $user['role'];

            if ($user['role'] === 'admin') {
                header("Location: dashboard.php");
            } elseif ($user['role'] === 'staff') {
                header("Location: staff_dashboard.php");
            }
            exit;
        }
    }

    $error = "Invalid username or password";
}

// Handle user registration
if (isset($_POST['register'])) {
    $reg_username = trim($_POST['reg_username']);
    $reg_password = trim($_POST['reg_password']);
    $reg_name = trim($_POST['reg_name']);
    $reg_age = intval($_POST['reg_age']);
    $reg_address = trim($_POST['reg_address']);
    $reg_role = trim($_POST['reg_role']);
    if ($reg_username && $reg_password && $reg_name && $reg_age && $reg_address && $reg_role) {
        $check = $conn->prepare("SELECT id FROM users WHERE username = ?");
        $check->bind_param('s', $reg_username);
        $check->execute();
        $check->store_result();
        if ($check->num_rows > 0) {
            $registerMsg = '<span style="color:red;">Username already exists.</span>';
        } else {
            // Store password as plain text (not secure)
            $stmt = $conn->prepare("INSERT INTO users (username, password, name, age, address, role, approved) VALUES (?, ?, ?, ?, ?, ?, 0)");
            $stmt->bind_param('sssiss', $reg_username, $reg_password, $reg_name, $reg_age, $reg_address, $reg_role);
            if ($stmt->execute()) {
                $registerMsg = '<span style="color:green;">Account created! You can now log in.</span>';
            } else {
                $registerMsg = '<span style="color:red;">Error creating account.</span>';
            }
            $stmt->close();
        }
        $check->close();
    } else {
        $registerMsg = '<span style="color:red;">Please fill in all fields.</span>';
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Jovern's Inventory Management System</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body class="login-body">

<div class="login-container">
    <div class="login-box">
        <div class="login-header">
            <h1 style="color:#fff;">Jovern's Fabricating Shop</h1>
            <p class="login-subtitle">Inventory Management System</p>
        </div>

        <div class="login-content">
            <?php if (!empty($error)): ?>
                <div class="login-error">
                    <span class="error-icon">⚠</span>
                    <span><?php echo $error; ?></span>
                </div>
            <?php endif; ?>
            <?php if (!empty($registerMsg)) echo $registerMsg; ?>
            <p class="login-greeting">Please enter your credentials to access the system</p>
            <form method="post" class="login-form" id="login-form">
                <div class="form-group">
                    <label for="username">Username</label>
                    <input type="text" id="username" name="username" placeholder="Enter your username" required>
                </div>
                <div class="form-group">
                    <label for="password">Password</label>
                    <input type="password" id="password" name="password" placeholder="Enter your password" required>
                </div>
                <button type="submit" name="login" class="login-btn">Sign In</button>
            </form>
            <div style="text-align:center;margin:15px 0;">
                <a href="#" id="show-signup" style="color:#666;text-decoration:underline;cursor:pointer;">Sign Up</a>
            </div>
            <div id="signup-form-container" style="display:none;">
                <hr style="margin:20px 0;">
                <h3>Create Account</h3>
                <form method="post" class="login-form">
                    <div class="form-group">
                        <label for="reg_name">Full Name</label>
                        <input type="text" id="reg_name" name="reg_name" placeholder="Enter your full name" required>
                    </div>
                    <div class="form-group">
                        <label for="reg_age">Age</label>
                        <input type="number" id="reg_age" name="reg_age" min="1" max="120" required>
                    </div>
                    <div class="form-group">
                        <label for="reg_address">Address</label>
                        <input type="text" id="reg_address" name="reg_address" placeholder="Enter your address" required>
                    </div>
                    <div class="form-group">
                        <label for="reg_username">Username</label>
                        <input type="text" id="reg_username" name="reg_username" placeholder="Choose a username" required>
                    </div>
                    <div class="form-group">
                        <label for="reg_password">Password</label>
                        <input type="password" id="reg_password" name="reg_password" placeholder="Choose a password" required>
                    </div>
                    <div class="form-group">
                        <label for="reg_role">Role</label>
                        <select id="reg_role" name="reg_role" required>
                            <option value="staff">Staff</option>
                            <option value="admin">Admin</option>
                        </select>
                    </div>
                    <button type="submit" name="register" class="login-btn" style="background:#666;">Create Account</button>
                </form>
                <div style="text-align:center;margin:10px 0;">
                    <a href="#" id="back-to-login" style="color:#666;text-decoration:underline;cursor:pointer;">Back to Login</a>
                </div>
            </div>
        </div>

        <div class="login-footer">
            <p>© 2026 Jovern’s Fabricating Shop | Built with Quality & Care
</p>
        </div>
    </div>
</div>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        var showSignup = document.getElementById('show-signup');
        var signupForm = document.getElementById('signup-form-container');
        var loginForm = document.getElementById('login-form');
        var backToLogin = document.getElementById('back-to-login');
        if (showSignup) {
            showSignup.onclick = function(e) {
                e.preventDefault();
                signupForm.style.display = 'block';
                loginForm.style.display = 'none';
                showSignup.style.display = 'none';
            };
        }
        if (backToLogin) {
            backToLogin.onclick = function(e) {
                e.preventDefault();
                signupForm.style.display = 'none';
                loginForm.style.display = 'block';
                showSignup.style.display = 'inline';
            };
        }
        // If there was a registration error or message, show the signup form automatically
        <?php if (!empty($registerMsg)) { ?>
            signupForm.style.display = 'block';
            loginForm.style.display = 'none';
            if (showSignup) showSignup.style.display = 'none';
        <?php } ?>
    });
</script>

</body>
</html>
