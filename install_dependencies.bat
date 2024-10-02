@echo off

REM Check if VCRUNTIME140.dll exists in the System32 folder
if exist "%SystemRoot%\System32\VCRUNTIME140.dll" (
    echo VCRUNTIME140.dll already exists. No need to install.
    goto end
)

REM Download the Visual C++ Redistributable if not already downloaded
echo Checking for Microsoft Visual C++ Redistributable...

if not exist vc_redist.x64.exe (
    echo Downloading Microsoft Visual C++ Redistributable package...
    powershell -Command "(New-Object System.Net.WebClient).DownloadFile('https://aka.ms/vs/17/release/vc_redist.x64.exe', 'vc_redist.x64.exe')"
)

REM Install the Redistributable silently
echo Installing Microsoft Visual C++ Redistributable package...
start /wait vc_redist.x64.exe /install /passive /norestart

REM Check if the installation was successful
if exist "%SystemRoot%\System32\VCRUNTIME140.dll" (
    echo Microsoft Visual C++ Redistributable successfully installed.
) else (
    echo Failed to install Microsoft Visual C++ Redistributable.
    exit /b 1
)

:end
echo All done! You can now run your Python executable.
pause
