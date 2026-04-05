@echo off
SET DIR=C:\Users\virat.arya\ETG\SoftsDatabase - Documents\Database\Hardmine\Non Fundamental\Cross Section
SET LOG=%DIR%\logs\automation_log.txt
SET PYTHON=C:\Users\virat.arya\.conda\envs\base\python.exe

echo. >> "%LOG%"
echo ============================================ >> "%LOG%"
echo Run started: %DATE% %TIME% >> "%LOG%"

echo Running Cross Section ingest... >> "%LOG%"
"%PYTHON%" "%DIR%\ingest.py" >> "%LOG%" 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: ingest.py failed >> "%LOG%"
    goto :error
)

echo Git add, commit and push... >> "%LOG%"
cd /d "%DIR%"
git add prices.parquet >> "%LOG%" 2>&1
git commit --allow-empty -m "Cross Section data update %DATE% %TIME%" >> "%LOG%" 2>&1
git push >> "%LOG%" 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Git push failed >> "%LOG%"
    goto :error
)

echo Run finished successfully: %DATE% %TIME% >> "%LOG%"
goto :end

:error
echo Sending failure email... >> "%LOG%"
powershell -Command "Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -UseSsl -From 'virat.arya@etgworld.com' -To 'virat.arya@etgworld.com' -Subject 'Cross Section Pipeline FAILED' -Body 'Check log: %LOG%' -Credential (New-Object PSCredential('virat.arya@etgworld.com', (ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force)))" >> "%LOG%" 2>&1

:end
