@echo off

color 04
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%" NEQ "0" (
	echo 请以管理员身份运行 %~nx0
	pause
	exit
)

color 02
set "batDir=%~dp0"
set "REGKEY=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
set "REGVALUE=Platform"
set "ARCH="
for /f "tokens=3" %%i in ('reg query "%REGKEY%" /v "%REGVALUE%" 2^>nul') do set "ARCH=%%i"
if "%ARCH%"=="x86" (
	echo Office 是 32 位版本。
	copy "%batDir%sppc32.dll" "%ProgramFiles%\Microsoft Office\root\vfs\System\sppc.dll"
	set "sppcDir=%windir%\SysWOW64\sppc.dll"
) else (
	echo Office 是 64 位版本。
	copy "%batDir%sppc64.dll" "%ProgramFiles%\Microsoft Office\root\vfs\System\sppc.dll"
	set "sppcDir=%windir%\System32\sppc.dll"
)
mklink "%ProgramFiles%\Microsoft Office\root\vfs\System\sppcs.dll" "%sppcDir%"
slmgr -ipk 2N382-D6PKK-QTX4D-2JJYK-M96P2
set "a=127.0.0.1 ols.officeapps.live.com"
set "file=C:\Windows\System32\drivers\etc\hosts"
(echo %a%) >> %file%
echo 已激活Microsoft 365。
pause