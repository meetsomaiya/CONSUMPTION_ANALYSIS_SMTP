@echo off
set "logfile=D:\xampp2\htdocs\LUBRICATION PORTAL SMTP\dispute_summary_report.log"
echo Script started at %date% %time% >> "%logfile%"

:: Run each PHP script and log the result, continue even if there is an error

echo Running: smtp_send_summary_report_dispute_new_one_manish.php >> "%logfile%"
"D:\xampp2\php\php-cgi.exe" -f "D:\xampp2\htdocs\LUBRICATION PORTAL SMTP\smtp_send_summary_report_dispute_new_one_manish.php" >> "%logfile%" 2>&1
if %errorlevel% neq 0 echo Failed: smtp_send_summary_report_dispute_new_one_manish.php >> "%logfile%"

