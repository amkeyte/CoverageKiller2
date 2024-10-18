@echo off
setlocal

rem Set the AutoRecover path
set autorecover_path=C:\Users\akeyte.PCM\AppData\Roaming\Microsoft\Word\

rem Delete all AutoRecover files (.asd)
if exist "%autorecover_path%*.asd" (
    echo Deleting AutoRecover files...
    del "%autorecover_path%*.asd" /Q
) else (
    echo No AutoRecover files found.
)

rem Delete temporary AutoRecover folders, excluding the STARTUP folder
if exist "%autorecover_path%*" (
    echo Deleting temporary AutoRecover folders...
    for /d %%i in ("%autorecover_path%*") do (
        if /i not "%%~nxi"=="STARTUP" (
            rmdir /s /q "%%i"
        )
    )
) else (
    echo No temporary folders found.
)

endlocal
