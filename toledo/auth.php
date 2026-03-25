<?php

require_once 'config.php';

function checkLogin() {
    if (!isset($_SESSION['user_id'])) {
        header('Location: index.php');
        exit;
    }
}


function checkAdmin() {
    checkLogin();
    if (!isset($_SESSION['role']) || $_SESSION['role'] !== 'admin') {
        header('Location: staff_dashboard.php'); 
        exit;
    }
}


function checkStaff() {
    checkLogin();
    if (!isset($_SESSION['role']) || $_SESSION['role'] !== 'staff') {
        header('Location: dashboard.php');
        exit;
    }
}

function currentUser($conn) {
    if (!isset($_SESSION['user_id'])) return null;
    $id = (int)$_SESSION['user_id'];
    $res = mysqli_query($conn, "SELECT * FROM users WHERE id=$id LIMIT 1");
    if ($res && mysqli_num_rows($res) === 1) {
        return mysqli_fetch_assoc($res);
    }
    return null;
}
