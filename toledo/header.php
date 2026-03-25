<?php
require_once 'config.php';
require_once 'db.php';
require_once 'auth.php'; 
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="styles.css">
    <title>bakeshop!</title>
</head>
<body>
<header class="sidebar">
    <div class="brand" style="font-size:2em;font-weight:bold;text-align:center;margin-bottom:10px;">Jovern's Fabricating Shop</div>
    <nav>
    <a href="dashboard.php"><span class="text">Dashboard</span></a>
    <a href="user_approvals.php"><span class="text">User Approvals</span></a>
    <a href="categories.php"><span class="text">Equipment Category</span></a>
    <a href="materials.php"><span class="text">Materials</span></a>
    <a href="products.php"><span class="text">Equipment</span></a>
    <a href="customers.php"><span class="text">Client</span></a>
    <a href="purchases.php"><span class="text">Equipment Purchase</span></a>
    <a href="reports.php"><span class="text">Sales Reports</span></a>
    <a href="archives.php"><span class="text">Archives</span></a>
    <a href="logout.php"><span class="text">Logout</span></a>
</nav>

</header>
<main class="container">

