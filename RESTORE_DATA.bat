
@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

:: Define MySQL credentials
SET MYSQL_DIR="F:\RMSSOFT310\MySQL_Py"
SET DB_USER=root
SET DB_NAME=sunildb
SET DB_PASSWORD=pass
SET SQL_FILE="F:\sunildb.sql"
SET DROP_FILE="drop_tables.sql"

:: Check if SQL file exists
IF NOT EXIST %SQL_FILE% (
    echo SQL file not found: %SQL_FILE%
    pause
    exit /b 1
)

:: Generate DROP TABLE statements
echo Generating DROP TABLE statements...
%MYSQL_DIR%\mysql -u %DB_USER% -p%DB_PASSWORD% -N -B -e "SELECT CONCAT('DROP TABLE IF EXISTS \`', table_name, '\`;') FROM information_schema.tables WHERE table_schema='%DB_NAME%';" > %DROP_FILE%

:: Execute DROP TABLE commands
echo Dropping tables from %DB_NAME% Please Wait ...
%MYSQL_DIR%\mysql -u %DB_USER% -p%DB_PASSWORD% %DB_NAME% < %DROP_FILE%
del %DROP_FILE%
echo All tables dropped.

:: Restore the database
echo Restoring database from %SQL_FILE% May Take Time, Please Wait...
%MYSQL_DIR%\mysql -u %DB_USER% -p%DB_PASSWORD% %DB_NAME% < %SQL_FILE%

IF %ERRORLEVEL% NEQ 0 (
    echo Restore failed!
    pause
    exit /b 1
)

echo Restore completed successfully!
pause

