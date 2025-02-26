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

function isValidEmail($email) {
    return filter_var($email, FILTER_VALIDATE_EMAIL);
}


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
    // $sqlIncharge = "
    //     SELECT DISTINCT [STATE], [AREA], [SITE],
    //            [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
    //     FROM [dbo].[site_area_incharge_mapping]";

    $sqlIncharge = "
    SELECT DISTINCT [STATE], [AREA], [SITE],
           [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO]
    FROM [dbo].[VW_site_area_incharge_mapping]";

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
        // $fcCount += fetchCount($db, $site, 'fc_oil_change_all_orders', 'FC_OIL_CHANGE ORDER', true);
        // $fcCount += fetchCount($db, $site, 'dispute_all_orders', 'FC_OIL_CHANGE ORDER', false);

        // $fcCount += fetchCount($db, $site, 'fc_oil_change', 'FC_OIL_CHANGE ORDER', true);
        // $fcCount += fetchCount($db, $site, 'dispute', 'FC_OIL_CHANGE ORDER', false);

        $fcCount += fetchCount($db, $site, 'fc_oil_change', 'FC_OIL_CHANGE ORDER', true);
        $fcCount += fetchCount($db, $site, 'dispute', 'FC_OIL_CHANGE ORDER', true);

        // Calculate GB-OIL CHANGE count
        // $gbCount += fetchCount($db, $site, 'gb_oil_change_all_orders', 'GB_OIL_CHANGE ORDER', true);
        // $gbCount += fetchCount($db, $site, 'dispute_all_orders', 'GB_OIL_CHANGE ORDER', false);

        // $gbCount += fetchCount($db, $site, 'gb_oil_change', 'GB_OIL_CHANGE ORDER', true);
        // $gbCount += fetchCount($db, $site, 'dispute', 'GB_OIL_CHANGE ORDER', false);

        
        $gbCount += fetchCount($db, $site, 'gb_oil_change', 'GB_OIL_CHANGE ORDER', true);
        $gbCount += fetchCount($db, $site, 'dispute', 'GB_OIL_CHANGE ORDER', true);

        // Calculate PD-OIL CHANGE count
        // $pdCount += fetchCount($db, $site, 'pd_oil_chg_order_all_orders', 'PD_OIL_CHG_ORDER', true);
        // $pdCount += fetchCount($db, $site, 'dispute_all_orders', 'PD_OIL_CHG_ORDER', false);

        // $pdCount += fetchCount($db, $site, 'PD_OIL_CHG_ORDER', 'PD_OIL_CHG_ORDER', true);
        // $pdCount += fetchCount($db, $site, 'dispute', 'PD_OIL_CHG_ORDER', false);

        $pdCount += fetchCount($db, $site, 'PD_OIL_CHG_ORDER', 'PD_OIL_CHG_ORDER', true);
        $pdCount += fetchCount($db, $site, 'dispute', 'PD_OIL_CHG_ORDER', true);

        // Calculate YD-OIL CHANGE count
        // $ydCount += fetchCount($db, $site, 'yd_oil_chg_order_all_orders', 'YD_OIL_CHG_ORDER', true);
        // $ydCount += fetchCount($db, $site, 'dispute_all_orders', 'YD_OIL_CHG_ORDER', false);

        // $ydCount += fetchCount($db, $site, 'YD_OIL_CHG_ORDER', 'YD_OIL_CHG_ORDER', true);
        // $ydCount += fetchCount($db, $site, 'dispute', 'YD_OIL_CHG_ORDER', false);

        $ydCount += fetchCount($db, $site, 'YD_OIL_CHG_ORDER', 'YD_OIL_CHG_ORDER', true);
        $ydCount += fetchCount($db, $site, 'dispute', 'YD_OIL_CHG_ORDER', true);

        // Calculate Grand Total
        $grandTotal = $fcCount + $gbCount + $pdCount + $ydCount;

            // Skip rows where Grand Total is 0
    if ($grandTotal === 0) {
        continue;
    }

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
    $fileName = 'Teco_Pending_Report.xlsx';
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
    $mail->addCC('meet.somaiya@suzlon.com');
    $mail->addCC('manish.jaiswal@suzlon.com');
    $mail->addCC('abhishek.devarkar@suzlon.com');

    // Add CCs for STATE ENGG HEAD, AREA INCHARGE, SITE INCHARGE, and STATE PMO
// foreach ($inchargeCombinations as $incharge) {
//     if (!empty($incharge['STATE ENGG HEAD'])) {
//         $mail->addAddress(trim($incharge['STATE ENGG HEAD']));
//     }
//     if (!empty($incharge['AREA INCHARGE'])) {
//         $mail->addAddress(trim($incharge['AREA INCHARGE']));
//     }
//     if (!empty($incharge['SITE INCHARGE'])) {
//         $mail->addAddress(trim($incharge['SITE INCHARGE']));
//     }
//     if (!empty($incharge['STATE PMO'])) {
//         $mail->addAddress(trim($incharge['STATE PMO']));
//     }
// }

foreach ($inchargeCombinations as $incharge) {
    if (!empty($incharge['STATE ENGG HEAD']) && isValidEmail($incharge['STATE ENGG HEAD'])) {
        $mail->addAddress(trim($incharge['STATE ENGG HEAD']));
    }
    if (!empty($incharge['AREA INCHARGE']) && isValidEmail($incharge['AREA INCHARGE'])) {
        $mail->addAddress(trim($incharge['AREA INCHARGE']));
    }
    if (!empty($incharge['SITE INCHARGE']) && isValidEmail($incharge['SITE INCHARGE'])) {
        $mail->addAddress(trim($incharge['SITE INCHARGE']));
    }
    if (!empty($incharge['STATE PMO']) && isValidEmail($incharge['STATE PMO'])) {
        $mail->addAddress(trim($incharge['STATE PMO']));
    }
}


    // Email content
    $mail->isHTML(true);
    // Email subject and body
   // Modify subject and body to include $currentDate
// $subject = "FW: GENTLE REMINDER 1 - PENDING TECO - IF PHYSICALLY GB FC YDPD OIL CHANGE WAS DONE - CLOSE IT BEFORE DT $currentDate WITHOUT FAIL (FC = 20 ,GB=14 ,YDPD =157)";
$subject = "FW: GENTLE REMINDER 1 - PENDING TECO - IF PHYSICALLY GB FC YDPD OIL CHANGE WAS DONE - CLOSE IT BEFORE DT $currentDate WITHOUT FAIL";
$body = "Respective AIC, SI & PMO,<br><br>
It has been observed that the physically oil change was done at location, but the SAP activity is pending due to that Actual status of oil change was not displayed in front of management.<br><br>
To avoid that please find the attached file containing suspect location where the TECO was pending.<br><br>
Kindly do the TECO of oil change order before $currentDate by end of the day if oil change was done & confirm on mail.<br><br>
NOTE – If you found an error during SAP – TECO process then write a mail with error snapshot 
to Mr Rahul Raut & Mr Harshvardhan at rahul.raut@suzlon.com & sbatech17@suzlon.com.<br><br>
1. GI DONE.<br>2. USED OIL RETURN TO SYSTEM<br>3. SAP OIL CHANGE PROCESS PENDING<br>4. GOODS MOVEMENT DONE";

    $mail->Subject = $subject;
    $mail->Body    = $body;

    // Attach the Excel file
    // $mail->addAttachment($fileName);

    // // Send the email
    // $mail->send();
    // echo "Email has been sent successfully.";

    // Retry mechanism for sending email
$emailSent = false;

while (!$emailSent) {
    try {
            // Attach the Excel file
    $mail->addAttachment($fileName);

        $mail->send();
        echo "Email sent successfully.";
        $emailSent = true; // Exit the loop when email is sent
    } catch (Exception $e) {
        echo "Email sending failed: {$mail->ErrorInfo}. Retrying...";
        sleep(5); // Wait for 5 seconds before retrying
    }
}

    // Step 5: Delete the Excel file
    unlink($fileName);
    echo "Excel file deleted successfully.";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>
