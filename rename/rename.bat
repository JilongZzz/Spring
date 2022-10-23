setlocal enabledelayedexpansion
 

for %%i in (*.xls) do (
    @REM echo f=%%n
    set f=%%i
    echo !f!
    echo !f:~0,13!.xls
    ren  "!f!" !f:~0,13!.xls
)

for %%i in (*.pdf) do (
    @REM echo f=%%n
    set f=%%i
    ren  "!f!" !f:~0,13!.pdf
)

for %%i in (*.zip) do (
    @REM echo f=%%n
    set f=%%i
    ren  "!f!" !f:~0,13!.zip
)

for %%i in (*.xlsx) do (
    @REM echo f=%%n
    set f=%%i
    ren  "!f!" !f:~0,13!.xlsx
)