@echo off
setlocal

echo ========================================
echo  TEAMPPT Add-in 설치
echo ========================================

set ADDIN_DIR=%~dp0
set DLL_PATH=%ADDIN_DIR%bin\Release\TeampptAddin.dll

if not exist "%DLL_PATH%" (
    set DLL_PATH=%ADDIN_DIR%bin\Debug\TeampptAddin.dll
)

if not exist "%DLL_PATH%" (
    echo [오류] TeampptAddin.dll을 찾을 수 없습니다.
    echo 먼저 Visual Studio에서 빌드하세요.
    pause
    exit /b 1
)

echo.
echo [1/2] COM 등록 중...
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "%DLL_PATH%" /codebase /tlb
if errorlevel 1 (
    echo.
    echo [오류] RegAsm 실패. 관리자 권한으로 실행하세요.
    pause
    exit /b 1
)

echo.
echo [2/2] PowerPoint Add-in 레지스트리 등록 중...
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect" /v "FriendlyName" /t REG_SZ /d "TEAMPPT" /f
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect" /v "Description" /t REG_SZ /d "TEAMPPT Header Assets Add-in" /f
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f

echo.
echo ========================================
echo  설치 완료! PowerPoint를 재시작하세요.
echo ========================================
pause
