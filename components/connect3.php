<?php
// Start session
// session_start();
if (session_status() == PHP_SESSION_NONE) {
    session_start();
}

// Database credentials
// $serverName = 'SGSMHPUN03784L\MSSQLSERVER01';
$serverName = 'SGHMHPUN03784L';
$database = 'Lubrication_Dashboard';
$username = 'meetsomaiya';
$password = 'Kitkat998';

// Establish database connection
$dsn = "sqlsrv:Server=$serverName;Database=$database;TrustServerCertificate=true";
$options = array(
    PDO::SQLSRV_ATTR_ENCODING => PDO::SQLSRV_ENCODING_UTF8,
    PDO::SQLSRV_ATTR_DIRECT_QUERY => true,
    PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION
);

try {
    $db = new PDO($dsn, $username, $password, $options);
} catch (PDOException $e) {
    echo 'Connection failed: ' . $e->getMessage();
    exit;
}
?>
