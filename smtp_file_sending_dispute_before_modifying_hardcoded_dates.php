<?php
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Include necessary libraries and PHPMailer autoload
require './vendor/autoload.php';
include './components/connect5.php'; // Include your PDO connection
require './PhpSpreadSheet/AnyFolder/PhpOffice/autoload.php'; // Load PhpSpreadsheet

$financialYear = $_GET['financial_year'] ?? 'FY 2024-2025';

$startDate = '2024-03-31';
$endDate = '2025-04-01';

// Step 1: Fetch existing order numbers from the dispute table
// $existingOrderQuery = "SELECT [Order No] FROM [dbo].[reason_for_dispute_and_pending_teco]";
$existingOrderQuery = "SELECT [Order No] FROM [dbo].[VW_reason_for_dispute_and_pending_teco]";
$existingOrderStmt = $db->query($existingOrderQuery);
$existingOrders = $existingOrderStmt->fetchAll(PDO::FETCH_COLUMN);

// Step 1: Generate Excel file

// Set timezone
date_default_timezone_set('Asia/Kolkata');

// Fetch today's date
$currentDate = date('Y-m-d');

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();

// // Fetch Site Area Incharge Mapping Information
// $siteMappingQuery = "SELECT [SITE], [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO] 
//     FROM [NewDatabase].[dbo].[site_area_incharge_mapping]";

// $siteMappingQuery = "SELECT DISTINCT [SITE], [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO], [extra] 
//     FROM [dbo].[site_area_incharge_mapping]";

// $siteMappingQuery = "SELECT DISTINCT [SITE], [STATE ENGG HEAD], [AREA INCHARGE], [SITE INCHARGE], [STATE PMO], [extra] 
//     FROM [dbo].[VW_site_area_incharge_mapping]";

// $siteMappingQuery = "SELECT DISTINCT [SITE], 
//                             [STATE ENGG HEAD], 
//                             [AREA INCHARGE], 
//                             [SITE INCHARGE], 
//                             [STATE PMO], 
//                             [extra],
//                             [MAINTENNACE INCHARGES],
//                             [GEAR BOX TEAM]
//                      FROM [NewDatabase].[dbo].[VW_site_area_incharge_mapping_new_format]";

$siteMappingQuery = "SELECT DISTINCT [SITE], 
[STATE ENGG HEAD], 
[AREA INCHARGE], 
[SITE INCHARGE], 
[STATE PMO], 
[extra],
[MAINTENNACE INCHARGES],
[GEAR BOX TEAM]
FROM [NewDatabase].[dbo].[VW_site_area_incharge_mapping]";

$siteMappingStmt = $db->query($siteMappingQuery);
$siteMappings = $siteMappingStmt->fetchAll(PDO::FETCH_ASSOC);
$siteMapping = [];
foreach ($siteMappings as $mapping) {
    $siteMapping[$mapping['SITE']] = $mapping;
}

$spreadsheet->getActiveSheet()->setTitle('Dispute');

// Fetch Pending Teco data from oil change and dispute tables
$pendingTecoTables = ['dispute_all_orders'];
$pendingTecoRows = [];

// Define the columns to fetch
$columnsToFetch = [
    '[Order No]',
    '[Function Loc]',
    '[Issue]',
    'TRY_CAST([Return] AS FLOAT) AS [Return]',  // Convert varchar to float
    '[Return Percentage]',
    '[Plant]',
    '[State]',
    '[Area]',
    '[Site]',
    '[Material]',
    '[Storage Location]',
    '[Move Type]',
    '[Material Document]',
    '[Description]',
    '[Val Type]',
    '[Posting Date]',
    '[Entry Date]',
    '[Quantity]',
    '[Order Type]',
    '[Component]',
    '[WTG Model]',
    '[Order]',
    '[Order Status]',
    '[Current Oil Change Date]'
];

// Loop through the Pending Teco tables
// Loop through the Pending Teco tables
foreach ($pendingTecoTables as $table) {
    // Prepare the column list for the SELECT statement
    $columnsList = implode(", ", $columnsToFetch);
    
    $query = "
    SELECT $columnsList FROM $table 
    WHERE [Posting Date] >= '2024-03-31' 
    AND [Posting Date] <= '2025-04-01'
    AND [date_of_insertion] = :currentDate";

    // Prepare and execute the query
    $stmt = $db->prepare($query);
    $stmt->bindParam(':currentDate', $currentDate);
    $stmt->execute();
    
    // Fetch rows from the current table
    $fetchedRows = $stmt->fetchAll(PDO::FETCH_ASSOC);
    
    // Filter out rows with existing order numbers
    foreach ($fetchedRows as $row) {
        if (!in_array($row['Order No'], $existingOrders)) {
            $pendingTecoRows[] = $row; // Only include rows that don't match existing orders
        }
    }
}

// Write the Pending Teco data into the sheet
$sheet = $spreadsheet->getActiveSheet();
if (!empty($pendingTecoRows)) {
    // Set the header row
    $headers = array_keys($pendingTecoRows[0]);
    
    // Add additional incharge headers
    $headers[] = 'STATE ENGG HEAD';
    $headers[] = 'AREA INCHARGE';
    $headers[] = 'SITE INCHARGE';
    $headers[] = 'STATE PMO';
    $headers[] = 'EXTRA';

    // Write headers to the first row
    foreach ($headers as $columnIndex => $header) {
        $sheet->setCellValueByColumnAndRow($columnIndex + 1, 1, $header);
    }

    // Populate data rows
// Step 1: Group data by unique combinations of Area Incharge, Site Incharge, State PMO, and new fields
$groupedData = [];
foreach ($pendingTecoRows as $row) {
    $stateEnggHead = $siteMapping[$row['Site']]['STATE ENGG HEAD'] ?? null;
    $areaIncharge = $siteMapping[$row['Site']]['AREA INCHARGE'] ?? null;
    $siteIncharge = $siteMapping[$row['Site']]['SITE INCHARGE'] ?? null;
    $statePmo = $siteMapping[$row['Site']]['STATE PMO'] ?? null;
    $extra = $siteMapping[$row['Site']]['extra'] ?? null;
    $maintenanceIncharges = $siteMapping[$row['Site']]['MAINTENNACE INCHARGES'] ?? null;
    $gearboxTeam = $siteMapping[$row['Site']]['GEAR BOX TEAM'] ?? null;

    $key = "{$areaIncharge}_{$siteIncharge}_{$statePmo}";

    if (!isset($groupedData[$key])) {
        $groupedData[$key] = [
            'rows' => [],
            'stateEnggHead' => $stateEnggHead,
            'areaIncharge' => $areaIncharge,
            'siteIncharge' => $siteIncharge,
            'statePmo' => $statePmo,
            'extra' => $extra,
            'maintenanceIncharges' => $maintenanceIncharges,
            'gearboxTeam' => $gearboxTeam,
        ];
    }
    $groupedData[$key]['rows'][] = $row;
}


// Step 2: For each group, create a separate Excel file and send an email
foreach ($groupedData as $group) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Dispute');

    // Set headers
    $headers = array_keys($pendingTecoRows[0]);
    // $headers = array_merge($headers, ['STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO', 'EXTRA']);
    $headers = array_merge(array_keys($pendingTecoRows[0]), ['STATE ENGG HEAD', 'AREA INCHARGE', 'SITE INCHARGE', 'STATE PMO', 'EXTRA', 'MAINTENANCE INCHARGES', 'GEARBOX TEAM']);
    foreach ($headers as $columnIndex => $header) {
        $sheet->setCellValueByColumnAndRow($columnIndex + 1, 1, $header);
    }

    // Populate data rows
    foreach ($group['rows'] as $rowIndex => $row) {
        $colIndex = 1;
        // foreach ($row as $key => $cellValue) {
        //     if ($key !== 'rn') {
        //         $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $cellValue);
        //     }
        // }
                // Write each field to the sheet
foreach ($row as $key => $cellValue) {
    if ($key !== 'rn') { // Skip 'rn' column from being written

        // Handle case where 'Order Type' is NULL or blank
        if ($key === 'Order Type' && (is_null($cellValue) || trim($cellValue) === '') && !empty($row['Order'])) {
            $cellValue = $row['Order']; // Set 'Order Type' to the value of 'Order'
        }

        // Check if the column is a date field and format accordingly
        if (in_array($key, ['Posting Date', 'Entry Date', 'Current Oil Change Date'])) {
            $formattedDate = \PhpOffice\PhpSpreadsheet\Shared\Date::stringToExcel(date('d-m-Y', strtotime($cellValue)));
            $sheet->setCellValueByColumnAndRow($colIndex, $rowIndex + 2, $formattedDate);
            $sheet->getStyleByColumnAndRow($colIndex, $rowIndex + 2)->getNumberFormat()->setFormatCode('DD/MM/YYYY');
        } else {
            $sheet->setCellValueByColumnAndRow($colIndex, $rowIndex + 2, $cellValue); // Start writing from the second row
        }
        $colIndex++; // Move to the next column
    }
}
        // Add incharge data
        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['stateEnggHead']);
        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['areaIncharge']);
        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['siteIncharge']);
        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['statePmo']);
        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['extra']);

        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['maintenanceIncharges']);
        $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex + 2, $group['gearboxTeam']);
    }

    // Save Excel file
    $filePath = "dispute_{$group['areaIncharge']}_{$group['siteIncharge']}_{$group['statePmo']}.xlsx";
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);

    // Send email
    $mail = new PHPMailer(true);
    try {
        // $mail->isSMTP();
        // $mail->Host = 'smtp.office365.com';
        // $mail->SMTPAuth = true;
        // $mail->Username = 'SVC_OMSApplications@suzlon.com';
        // $mail->Password = 'Suzlon@123';
        // $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
        // $mail->Port = 587;

        $mail->isSMTP();
        $mail->Host = 'smtp.office365.com';
        $mail->SMTPAuth = true;
        $mail->Username = 'svc_fleetM@suzlon.com';
        $mail->Password = 'Suzlon@123';
        $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
        $mail->Port = 587;

        // Sender and recipients
        // $mail->setFrom('SVC_OMSApplications@suzlon.com', 'Suzlon OMS Applications');
        $mail->setFrom('svc_fleetM@suzlon.com', 'Suzlon OMS Applications');
        // $mail->addAddress($group['areaIncharge']);
        // $mail->addAddress($group['siteIncharge']);
        // $mail->addAddress($group['statePmo']);
        // $mail->addAddress($group['extra']);

// Add addresses only if they are set and not empty
if (!empty($group['areaIncharge'])) {
    $mail->addAddress($group['areaIncharge']);
}

if (!empty($group['siteIncharge'])) {
    $mail->addAddress($group['siteIncharge']);
}

if (!empty($group['statePmo'])) {
    $mail->addAddress($group['statePmo']);
}

// Add extra address only if it's set and not empty
if (!empty($group['extra'])) {
    $mail->addAddress($group['extra']);
}

if (!empty($group['maintenanceIncharges'])) $mail->addAddress($group['maintenanceIncharges']);
if (!empty($group['gearboxTeam'])) $mail->addAddress($group['gearboxTeam']);


        // $mail->addAddress('meetsomaiya5@gmail.com');
        $mail->addCC('meet.somaiya@suzlon.com');
        $mail->addCC('manish.jaiswal@suzlon.com');
        $mail->addCC('abhishek.devarkar@suzlon.com');

     // Email subject and body
    // $mail->Subject = 'DO TECO PENDING GB FC YDPD OIL CHANGE ORDER IMMEDAITELY ';
    $mail->Subject = 'REQUIRED SITE JUSTIFICATION AGAINST THE OIL CHANGE ORDER OF 80% LESS OIL RETURN';
    // $mail->Body = "Respective AIC, SI & PMO,\n\nPlease find the attached file that contains suspect locations where the physical oil change was done but TECO is pending.\nKindly complete the TECO of the oil change order by the end of today if the physical oil change was done.\n\nNOTE – If you encounter an error during the SAP-TECO process, please send an email with the error snapshot to Mr. Rahul Raut (rahul.raut@suzlon.com) & Mr. Harshvardhan (sbatech17@suzlon.com).\n\n1. GI DONE.\n2. USED OIL RETURNED TO SYSTEM.\n3. SAP OIL CHANGE PROCESS PENDING.\n4. GOODS MOVEMENT DONE.";

    // $mail->Body = "Respective AIC, SI & PMO,\n\nPlease find the attached file that contains suspect locations where the physical oil change was done but TECO is pending.\nKindly complete the TECO of the oil change order by the end of today if the physical oil change was done.\n\nPlease refer to SOP on this given link ==> http://10.102.0.192:778/admins/sop_for_consumption_analysis.php\n\nNOTE – If you encounter an error during the SAP-TECO process, please send an email with the error snapshot to Mr. Rahul Raut (rahul.raut@suzlon.com) & Mr. Harshvardhan (sbatech17@suzlon.com).\n\n1. GI DONE.\n2. USED OIL RETURNED TO SYSTEM.\n3. SAP OIL CHANGE PROCESS PENDING.\n4. GOODS MOVEMENT DONE.";

    // $mail->Body = "Respective AIC, SIC, PMO,\n\nPlease find the attached file contains the consumption details of Gear box, Fluid coupling, Yaw & Pitch drive oil till date.\nKindly provide site justification of oil change order where the oil return is less than 80%, only oil issue, only oil return by end of the day today.";

    // $mail->Body = "Respective AIC, SIC, PMO,\n\nPlease find the attached file containing the consumption details of Gear box, Fluid coupling, Yaw & Pitch drive oil till date. Kindly provide site justification of oil change order where the oil return is less than 80%, only oil issue, only oil return by end of the day today.\n\nPlease refer to SOP on this given link ==> https://suzoms.suzlon.com/LubricationDashboard/admins/sop_for_consumption_analysis2.php";

    // previous before fleet manager integration   
    // $mail->Body = "Respective AIC, SIC, PMO,\n\nPlease find the attached file containing the consumption details of Gear box, Fluid coupling, Yaw & Pitch drive oil till date. Kindly provide site justification of oil change order where the oil return is less than 80%, only oil issue, only oil return by end of the day today.\n\nPlease refer to SOP on this given link ==> https://suzoms.suzlon.com/LubricationDashboard/admins/sop_for_consumption_analysis2.php\n\nFor application details, please visit: https://suzoms.suzlon.com/LubricationDashboard/login/login.php";

        $mail->Body = "Respective AIC, SIC, PMO,\n\nPlease find the attached file containing the consumption details of Gear box, Fluid coupling, Yaw & Pitch drive oil till date. Kindly provide site justification of oil change order where the oil return is less than 80%, only oil issue, only oil return by end of the day today.\n\nPlease refer to SOP on this given link ==> https://suzoms.suzlon.com/LubricationPortal/index.html#/sop-consumption_user\n\nFor application details, please visit: https://suzoms.suzlon.com/FleetM/#/signin";


    // $mail->Body = "Respective AIC, SI & PMO,\n\nPlease find the attached file that contains suspect locations where the physical oil change was done but TECO is pending.\nKindly complete the TECO of the oil change order by the end of today if the physical oil change was done.\n\nPlease refer to SOP on this given link ==> http://10.102.0.192:778/admins/sop_for_consumption_analysis.php\n\nNOTE  If you encounter an error during the SAP-TECO process, please send an email with the error snapshot to Mr. Rahul Raut (rahul.raut@suzlon.com) & Mr. Harshvardhan (sbatech17@suzlon.com).\n\n1. GI DONE.\n2. USED OIL RETURNED TO SYSTEM.\n3. SAP OIL CHANGE PROCESS PENDING.\n4. GOODS MOVEMENT DONE.";

        // Attach Excel file
        $mail->addAttachment($filePath);

        // $mail->send();
           // Retry mechanism for sending email
$emailSent = false;

while (!$emailSent) {
    try {
        $mail->send();
        echo "Email sent successfully.";
        $emailSent = true; // Exit the loop when email is sent
    } catch (Exception $e) {
        echo "Email sending failed: {$mail->ErrorInfo}. Retrying...";
        sleep(5); // Wait for 5 seconds before retrying
    }
}

    } catch (Exception $e) {
        echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
    }

    // Optionally, delete the file after sending if not needed on server
    unlink($filePath);
}

}

?>
