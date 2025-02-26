<?php
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

require './vendor/autoload.php';
include './components/connect5.php';
require './PhpSpreadSheet/AnyFolder/PhpOffice/autoload.php'; // Load PhpSpreadsheet

// Initialize Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle("Oil Change Summary Report");

// Set header row
$headers = [
    'STATE', 'AREA', 'SITE', 'STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO',
    'FC-OIL CHANGE', 'GB-OIL CHANGE', 'PD-OIL CHANGE', 'YD-OIL CHANGE', 'Grand Total'
];
$sheet->fromArray($headers, NULL, 'A1');

$row = 2; // Starting row for data entries

// Helper function for fetching counts from a specific table
function fetchCount($db, $site, $table, $orderType, $useOrderStatus = true) {
    $sql = "
        SELECT COUNT(*) AS order_count
        FROM $table
        WHERE [Site] = :site AND [Order] = :orderType";
    if ($useOrderStatus) {
        $sql .= " AND ([Order Status] = 'released' OR [Order Status] = 'in process')";
    }
    $stmt = $db->prepare($sql);
    $stmt->execute([':site' => $site, ':orderType' => $orderType]);
    return $stmt->fetchColumn();
}

try {
    // Step 1: Fetch distinct incharge combinations
    $sqlIncharge = "
        SELECT DISTINCT [STATE], [AREA], [SITE],
               [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
        FROM [NewDatabase].[dbo].[site_area_incharge_mapping]";
    $stmtIncharge = $db->prepare($sqlIncharge);
    $stmtIncharge->execute();
    $inchargeCombinations = $stmtIncharge->fetchAll(PDO::FETCH_ASSOC);

    // Step 2: Process each incharge combination
    foreach ($inchargeCombinations as $incharge) {
        $state = $incharge['STATE'];
        $area = $incharge['AREA'];
        $site = $incharge['SITE'];

        // Initialize counts for each oil change type
        $fcCount = 0;
        $gbCount = 0;
        $pdCount = 0;
        $ydCount = 0;

        // Calculate FC-OIL CHANGE count
        // $fcCount += fetchCount($db, $site, 'fc_oil_change_all_orders', 'FC_OIL_CHANGE', true);
        // $fcCount += fetchCount($db, $site, 'dispute_all_orders', 'FC_OIL_CHANGE', false);

        $fcCount += fetchCount($db, $site, 'fc_oil_change_all_orders', 'FC_OIL_CHANGE ORDER', true);
        $fcCount += fetchCount($db, $site, 'dispute_all_orders', 'FC_OIL_CHANGE ORDER', false);

        // Calculate GB-OIL CHANGE count
        // $gbCount += fetchCount($db, $site, 'gb_oil_change_all_orders', 'GB_OIL_CHANGE', true);
        // $gbCount += fetchCount($db, $site, 'dispute_all_orders', 'GB_OIL_CHANGE', false);

        $gbCount += fetchCount($db, $site, 'gb_oil_change_all_orders', 'GB_OIL_CHANGE ORDER', true);
        $gbCount += fetchCount($db, $site, 'dispute_all_orders', 'GB_OIL_CHANGE ORDER', false);

        // Calculate PD-OIL CHANGE count
        $pdCount += fetchCount($db, $site, 'pd_oil_chg_order_all_orders', 'PD_OIL_CHG_ORDER', true);
        $pdCount += fetchCount($db, $site, 'dispute_all_orders', 'PD_OIL_CHG_ORDER', false);

        // Calculate YD-OIL CHANGE count
        $ydCount += fetchCount($db, $site, 'yd_oil_chg_order_all_orders', 'YD_OIL_CHG_ORDER', true);
        $ydCount += fetchCount($db, $site, 'dispute_all_orders', 'YD_OIL_CHG_ORDER', false);

        // Calculate Grand Total
        $grandTotal = $fcCount + $gbCount + $pdCount + $ydCount;

        // Step 3: Populate Excel row
        $sheet->setCellValue("A{$row}", $state);
        $sheet->setCellValue("B{$row}", $area);
        $sheet->setCellValue("C{$row}", $site);
        $sheet->setCellValue("D{$row}", $incharge['STATE ENGG HEAD']);
        $sheet->setCellValue("E{$row}", $incharge['AREA INCHARGE']);
        $sheet->setCellValue("F{$row}", $incharge['SITE INCHARGE']);
        $sheet->setCellValue("G{$row}", $incharge['STATE PMO']);
        $sheet->setCellValue("H{$row}", $fcCount);
        $sheet->setCellValue("I{$row}", $gbCount);
        $sheet->setCellValue("J{$row}", $pdCount);
        $sheet->setCellValue("K{$row}", $ydCount);
        $sheet->setCellValue("L{$row}", $grandTotal);

        $row++; // Move to the next row
    }

    // Save Excel file
    $writer = new Xlsx($spreadsheet);
    $writer->save('Oil_Change_Report.xlsx');
    echo "Excel report created successfully.";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>
