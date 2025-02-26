@echo off
set "logfile=D:\xampp2\htdocs\LUBRICATION PORTAL SMTP\teco_pending.log"
echo Script started at %date% %time% >> "%logfile%"

:: Run each PHP script and log the result, continue even if there is an error

echo Running: smtp_file_sending_pending_teco.php >> "%logfile%"
"D:\xampp2\php\php-cgi.exe" -f "D:\xampp2\htdocs\LUBRICATION PORTAL SMTP\smtp_file_sending_pending_teco.php" >> "%logfile%" 2>&1
if %errorlevel% neq 0 echo Failed: smtp_file_sending_pending_teco.php >> "%logfile%"

