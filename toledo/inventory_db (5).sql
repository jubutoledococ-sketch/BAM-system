CREATE DATABASE IF NOT EXISTS inventory_db;
USE inventory_db;

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";
SET NAMES utf8mb4;

CREATE TABLE `categories` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `category_name` varchar(100) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `customers` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `customer_name` varchar(100) NOT NULL,
  `contact` varchar(100),
  `email` varchar(255),
  `address` varchar(255),
  `client_type` varchar(50) DEFAULT 'Walk-in',
  `payment_terms` varchar(50) DEFAULT 'COD',
  `created_at` date NOT NULL DEFAULT curdate(),
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `materials` (
  `material_id` int(11) NOT NULL AUTO_INCREMENT,
  `material_name` varchar(150) NOT NULL,
  `supplier` varchar(150) NOT NULL,
  `price` decimal(10,2) NOT NULL,
  `image` varchar(255),
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  PRIMARY KEY (`material_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `products` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(100) NOT NULL,
  `category_id` int(11),
  `price` decimal(10,2) NOT NULL,
  `original_price` decimal(10,2),
  `stock` int(11) NOT NULL DEFAULT 0,
  `labor_cost` decimal(10,2) DEFAULT 0.00,
  `material_cost` decimal(10,2) DEFAULT 0.00,
  `image` varchar(255),
  PRIMARY KEY (`id`),
  KEY `category_id` (`category_id`),
  CONSTRAINT `products_ibfk_1`
    FOREIGN KEY (`category_id`) REFERENCES `categories` (`id`)
    ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `product_materials` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `product_id` int(11) NOT NULL,
  `material_id` int(11) NOT NULL,
  `qty` int(11) NOT NULL,
  PRIMARY KEY (`id`),
  KEY `product_id` (`product_id`),
  KEY `material_id` (`material_id`),
  CONSTRAINT `product_materials_ibfk_1`
    FOREIGN KEY (`product_id`) REFERENCES `products` (`id`) ON DELETE CASCADE,
  CONSTRAINT `product_materials_ibfk_2`
    FOREIGN KEY (`material_id`) REFERENCES `materials` (`material_id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `sales` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `product_id` int(11) NOT NULL,
  `quantity` int(11) NOT NULL,
  `total_price` decimal(10,2) NOT NULL,
  `date` date NOT NULL,
  `purchase_id` int(11),
  PRIMARY KEY (`id`),
  KEY `product_id` (`product_id`),
  CONSTRAINT `sales_ibfk_1`
    FOREIGN KEY (`product_id`) REFERENCES `products` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `payments` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `sale_id` int(11) NOT NULL,
  `payment_method` varchar(50) NOT NULL,
  `amount` decimal(10,2) NOT NULL,
  `date` date NOT NULL,
  PRIMARY KEY (`id`),
  KEY `sale_id` (`sale_id`),
  CONSTRAINT `payments_ibfk_1`
    FOREIGN KEY (`sale_id`) REFERENCES `sales` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `purchases` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `product_id` int(11) NOT NULL,
  `customer_id` int(11),
  `quantity` int(11) NOT NULL,
  `cost` decimal(10,2) NOT NULL,
  `date` date NOT NULL,
  `status` enum('Pending','Fabricating','Quality Check','Delivered','Repair') DEFAULT 'Pending',
  `handled_by` int(11),
  `expected_delivery` date,
  `notes` text,
  `completed_date` date,
  `original_cost` decimal(10,2),
  `discount` decimal(10,2) DEFAULT 0.00,
  `created_by` varchar(20),
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `stock_logs` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `product_id` int(11) NOT NULL,
  `change_type` varchar(20) NOT NULL,
  `quantity` int(11) NOT NULL,
  `date` date NOT NULL,
  PRIMARY KEY (`id`),
  KEY `product_id` (`product_id`),
  CONSTRAINT `stock_logs_ibfk_1`
    FOREIGN KEY (`product_id`) REFERENCES `products` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `suppliers` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `supplier_name` varchar(100) NOT NULL,
  `contact` varchar(100),
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `users` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `username` varchar(50),
  `password` varchar(255),
  `name` varchar(100),
  `age` int(3),
  `address` varchar(255),
  `role` varchar(20),
  `approved` tinyint(1) DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

INSERT INTO `users` (`id`, `username`, `password`, `role`) VALUES
(1, 'admin', 'admin123', 'admin'),
(2, 'staff', 'staff123', 'staff');
COMMIT;
