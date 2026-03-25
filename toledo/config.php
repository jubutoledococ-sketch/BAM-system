<?php
// config.php

// Database credentials
$host = 'localhost';
$db   = 'inventory_db';
$user = 'root';
$pass = '';

// Create connection
$conn = mysqli_connect($host, $user, $pass, $db);

// Check connection
if (!$conn) {
    die('Database connection failed: ' . mysqli_connect_error());
}

// Start session if not started
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}

// Optional: set default timezone
date_default_timezone_set('Asia/Manila');

// Optional: set mysqli charset to UTF-8
mysqli_set_charset($conn, 'utf8mb4');
