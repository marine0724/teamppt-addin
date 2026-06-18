@echo off
setlocal

echo ========================================
echo  TEAMPPT Add-in 제거
echo ========================================

set ADDIN_DIR=%~dp0
set DLL_PATH=%ADDIN_DIR%bin\Release\TeampptAddin.dll

if not exist "%DLL_PATH%" (
    set DLL_PATH=%ADDIN_DIR%bin\Debug\TeampptAddin.dll
)

echo.
echo [1/2] COM 등록 해제 중...
if exist "%DLL_PATH%" (
    %WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "%DLL_PATH%" /unregister
)

echo.
echo [2/2] PowerPoint Add-in 레지스트리 제거 중...
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect" /f 2>nul

echo.
echo ========================================
echo  제거 완료!
echo ========================================
pause
