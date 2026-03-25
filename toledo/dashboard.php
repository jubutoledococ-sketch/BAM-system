<?php
require_once 'header.php';
require_once 'db.php';
// require_once 'sidebar.php'; // Sidebar removed to restore original layout

// Handle staff account creation
$staffMsg = '';
if (isset($_POST['create_staff'])) {
    $username = trim($_POST['staff_username']);
    $password = trim($_POST['staff_password']);
    if ($username && $password) {
        // Check if username already exists
        $check = $conn->prepare("SELECT id FROM users WHERE username = ?");
        $check->bind_param('s', $username);
        $check->execute();
        $check->store_result();
        if ($check->num_rows > 0) {
            $staffMsg = '<span style="color:red;">Username already exists.</span>';
        } else {
            // Store password as plain text (not secure)
            $stmt = $conn->prepare("INSERT INTO users (username, password, role) VALUES (?, ?, 'staff')");
            $stmt->bind_param('ss', $username, $password);
            if ($stmt->execute()) {
                $staffMsg = '<span style="color:green;">Staff account created successfully!</span>';
            } else {
                $staffMsg = '<span style="color:red;">Error creating account.</span>';
            }
            $stmt->close();
        }
        $check->close();
    } else {
        $staffMsg = '<span style="color:red;">Please fill in all fields.</span>';
    }
}

$dailySales = mysqli_fetch_assoc(mysqli_query($conn, "SELECT IFNULL(SUM(total_price),0) AS amt FROM sales WHERE date = CURDATE()"));
$monthlySales = mysqli_fetch_assoc(mysqli_query($conn, "SELECT IFNULL(SUM(total_price),0) AS amt FROM sales WHERE DATE_FORMAT(date,'%Y-%m') = DATE_FORMAT(CURDATE(),'%Y-%m')"));
$totalRevenue = mysqli_fetch_assoc(mysqli_query($conn, "SELECT IFNULL(SUM(total_price),0) AS amt FROM sales"));
$totalClients = mysqli_fetch_assoc(mysqli_query($conn, "SELECT COUNT(*) AS cnt FROM customers"));
$lowStock = mysqli_query($conn, "SELECT name, stock FROM products WHERE stock < 5 ORDER BY stock ASC");
$recentSales = mysqli_query($conn, "SELECT id, total_price, date 
                                   FROM sales 
                                   ORDER BY date DESC 
                                   LIMIT 5");

// Sales Trend Data - Last 30 days
$salesTrendQuery = mysqli_query($conn, 
    "SELECT DATE(date) as sales_date, IFNULL(SUM(total_price),0) as daily_total
     FROM sales
     WHERE date >= DATE_SUB(CURDATE(), INTERVAL 30 DAY)
     GROUP BY DATE(date)
     ORDER BY sales_date ASC"
);

$trendDates = [];
$trendValues = [];
while($row = mysqli_fetch_assoc($salesTrendQuery)) {
    $trendDates[] = date('M d', strtotime($row['sales_date']));
    $trendValues[] = (float)$row['daily_total'];
}

// Sales by Category - Last 30 days
$categoryTrendQuery = mysqli_query($conn,
    "SELECT c.category_name, IFNULL(SUM(s.total_price),0) as category_total
     FROM sales s
     JOIN products p ON s.product_id = p.id
     JOIN categories c ON p.category_id = c.id
     WHERE s.date >= DATE_SUB(CURDATE(), INTERVAL 30 DAY)
     GROUP BY c.id, c.category_name
     ORDER BY category_total DESC"
);

$categoryNames = [];
$categoryValues = [];
while($row = mysqli_fetch_assoc($categoryTrendQuery)) {
    $categoryNames[] = $row['category_name'];
    $categoryValues[] = (float)$row['category_total'];
}

?>

<h1>Welcom Admin!</h1>

<!-- Metrics Cards -->
<div class="grid">
    <div class="card" style="background-color:#8c8c8c;color:white;">
        <h3>Daily Sales</h3>
        <p>₱<?php echo number_format($dailySales['amt'], 2); ?></p>
    </div>
    <div class="card" style="background-color:#b1a07d;color:white;">
        <h3>Monthly Sales</h3>
        <p>₱<?php echo number_format($monthlySales['amt'], 2); ?></p>
    </div>
    <div class="card" style="background-color:#7f8c8d;color:white;">
        <h3>Total Revenue</h3>
        <p>₱<?php echo number_format($totalRevenue['amt'], 2); ?></p>
    </div>
    <div class="card" style="background-color:#666666;color:white;">
        <h3>Total Clients</h3>
        <p><?php echo (int)$totalClients['cnt']; ?></p>
    </div>
</div>

<h3>Recent Sales</h3>
<table>
    <tr>
        <th>Total Price</th>
        <th>Date</th>
    </tr>
 <?php while($sale = mysqli_fetch_assoc($recentSales)): ?>
    <tr>
        <td>₱<?php echo number_format($sale['total_price'], 2); ?></td>
        <td><?php echo $sale['date']; ?></td>
    </tr>
<?php endwhile; ?>
</table>

<!-- Placeholder for Charts -->
<h3 style="margin-top: 10px;">Sales Trends (Last 30 Days)</h3>
<div class="chart-container" style="position: relative; width: 100%; height: 350px; margin: 20px 0; background: white; padding: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
    <canvas id="salesTrendChart"></canvas>
</div>

<!-- Chart.js Library -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script>
    // Sales Trend Chart
    const trendCtx = document.getElementById('salesTrendChart').getContext('2d');
    const salesTrendChart = new Chart(trendCtx, {
        type: 'line',
        data: {
            labels: <?php echo json_encode($trendDates); ?>,
            datasets: [{
                label: 'Daily Sales (₱)',
                data: <?php echo json_encode($trendValues); ?>,
                borderColor: '#667eea',
                backgroundColor: 'rgba(102, 126, 234, 0.1)',
                borderWidth: 3,
                fill: true,
                tension: 0.4,
                pointRadius: 5,
                pointBackgroundColor: '#667eea',
                pointBorderColor: '#ffffff',
                pointBorderWidth: 2,
                pointHoverRadius: 7,
                pointHoverBackgroundColor: '#764ba2',
                hoverBackgroundColor: 'rgba(102, 126, 234, 0.2)'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: { size: 13, weight: '500' },
                        color: '#333333',
                        padding: 15
                    }
                },
                title: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)',
                        drawBorder: false
                    },
                    ticks: {
                        color: '#999999',
                        font: { size: 12 },
                        callback: function(value) {
                            return '₱' + value.toLocaleString();
                        }
                    }
                },
                x: {
                    grid: {
                        display: false,
                        drawBorder: false
                    },
                    ticks: {
                        color: '#999999',
                        font: { size: 12 }
                    }
                }
            }
        }
    });
</script>

<?php require_once 'footer.php'; ?>