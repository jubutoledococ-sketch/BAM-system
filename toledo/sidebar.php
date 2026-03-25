<?php
require_once 'config.php';
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
    <a href="staff_dashboard.php"><span class="text">Dashboard</span></a>
    <a href="staff_products.php"><span class="text">Equipment</span></a>
    <a href="staff_customer.php"><span class="text">Client</span></a>
    <a href="staff_purchases.php"><span class="text">Equipment Purchase</span></a>
    <a href="logout.php"><span class="text">Logout</span></a>
</nav>

</header>
<main class="container">

