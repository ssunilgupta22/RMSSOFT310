
    @echo off

    echo %time%
    timeout 5 > NUL
    echo %time%

    move  F:\RMSSOFT310\mrms_pro.exe F:\RMSSOFT310\win_rms\RMS_BACKUP 
    copy F:\RMSSOFT310\MyfileDownload\mrms_pro.exe F:\RMSSOFT310\mrms_pro.exe 
    start F:\RMSSOFT310\mrms_pro.exe 
   exit /b