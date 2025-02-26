<?php
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

// Set timezone
date_default_timezone_set('Asia/Kolkata');

// Fetch today's date
$currentDate = date('Y-m-d');

// Calculate the current and previous months
$currentMonth = date('F Y'); // e.g., "September 2024"
$previousMonth = (new DateTime('first day of last month'))->format('F Y'); // e.g., "August 2024"

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

// Helper function for fetching counts from the dispute table
function fetchDisputeCount($db, $site, $orderType) {
    $sql = "
        SELECT COUNT(*) AS order_count
        FROM dispute
        WHERE [Site] = :site AND [Order] = :orderType";
    $stmt = $db->prepare($sql);
    $stmt->execute([':site' => $site, ':orderType' => $orderType]);
    return $stmt->fetchColumn();
}

try {
    // Variable to hold the total of all grand totals
    $totalGrandOrders = 0;

    // Step 1: Fetch distinct incharge combinations
    $sqlIncharge = "
        SELECT DISTINCT [STATE], [AREA], [SITE],
               [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
        FROM [dbo].[site_area_incharge_mapping]";
    $stmtIncharge = $db->prepare($sqlIncharge);
    $stmtIncharge->execute();
    $inchargeCombinations = $stmtIncharge->fetchAll(PDO::FETCH_ASSOC);

    // Step 2: Process each incharge combination
    foreach ($inchargeCombinations as $incharge) {
        $state = $incharge['STATE'];
        $area = $incharge['AREA'];
        $site = $incharge['SITE'];

        // Initialize counts for each oil change type
        $fcCount = fetchDisputeCount($db, $site, 'FC_OIL_CHANGE ORDER');
        $gbCount = fetchDisputeCount($db, $site, 'GB_OIL_CHANGE ORDER');
        $pdCount = fetchDisputeCount($db, $site, 'PD_OIL_CHG_ORDER');
        $ydCount = fetchDisputeCount($db, $site, 'YD_OIL_CHG_ORDER');

        // Calculate Grand Total
        $grandTotal = $fcCount + $gbCount + $pdCount + $ydCount;

        // Add to the total of all grand totals
        $totalGrandOrders += $grandTotal;

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
    $fileName = 'Dispute_Report.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save($fileName);
    echo "Excel report created successfully.";

    // Step 4: Send email with the Excel file attached
    $mail = new PHPMailer(true);
    $mail->isSMTP();
    $mail->Host = 'smtp.office365.com';
    $mail->SMTPAuth = true;
    $mail->Username = 'SVC_OMSApplications@suzlon.com';
    $mail->Password = 'Suzlon@123';
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
    $mail->Port = 587;

    // Sender and recipients
    $mail->setFrom('SVC_OMSApplications@suzlon.com', 'Suzlon OMS Applications');
    $mail->addAddress('meet.somaiya@suzlon.com');
    // $mail->addAddress('abhishek.devarkar@suzlon.com');

    // Email content
    $mail->isHTML(true);

    // Use the dynamic total in the subject
    $subject = "CONSUMPTION ANALYSIS FY 24-25 - GB FC TILL {$previousMonth} - CLOSE IT BEFORE 13 {$currentMonth} - SITE FEEDBACK PENDING {$totalGrandOrders} CONSUMPTION ORDER";
    $body = "
    Respective AIC, SIC, PMO,<br><br>
    Please find the attached file contains the consumption details of Gear box, Fluid coupling oil till {$previousMonth}.<br><br>
    Kindly provide site justification of order where the oil return is less than 80%, only oil issue, only oil return before <strong>13 {$currentMonth}</strong>.<br><br>
    This is to inform you that I have highlighted the summary of GB FC oil issue vs oil return as per oil change order for the current financial year 24-25.<br><br>
    The attached file contains order-wise details available with movement type oil issue & oil return for GB, FC.<br>
    Those states that have not achieved the oil return target kindly provide the order-wise justification.<br><br>
    <strong>ACTION TAKEN BY SITE</strong><br>
    Please check the data of your state, site, and location by clicking on the number of b/d or oil change in the summary sheet.<br>
    Kindly check the location-wise order & do the correction by doing TECO of Oil change order and Breakdown order.<br>
    If oil change was done and TECO is pending, ensure the oil change date is reflected in SAP oil change history or not.<br>
    Check the quantity of oil issue VS oil return also and if the return oil target is not achieved, provide the justification.<br><br>
    Order-wise details are divided into three categories:<br>
    1. Oil return Target achieved.<br>
    2. Oil return Target not achieved.<br>
    3. Only oil return without issue.<br><br>
    <strong>NOTE</strong>: Site representatives must provide justification against the return oil missing query in the <strong>SITE FEEDBACK PENDING COLUMN</strong>.<br>";

    $mail->Subject = $subject;
    $mail->Body    = $body;

    // Attach the Excel file
    $mail->addAttachment($fileName);

    // Send the email
    $mail->send();
    echo "Email has been sent successfully.";

    // Step 5: Delete the Excel file
    unlink($fileName);
    echo "Excel file deleted successfully.";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>
