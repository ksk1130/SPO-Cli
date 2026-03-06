@echo off
REM build-exe.bat [net48|net10]
REM net48: .NET Framework 4.8向け（ランタイム同梱なし）
REM net10: .NET 10向け self-contained 単一EXE（ランタイム同梱）

setlocal enabledelayedexpansion

set TARGET=%~1
if "%TARGET%"=="" set TARGET=net10

if /I "%TARGET%"=="net48" goto build_net48
if /I "%TARGET%"=="net10" goto build_net10
if /I "%TARGET%"=="net10.0" goto build_net10

echo [ERROR] Unknown target: %TARGET%
echo Usage: build-exe.bat [net48^|net10]
exit /b 1

:build_net48
echo Building net48 EXE (framework-dependent)...
dotnet build SpoCli.csproj -c Release -f net48
if %ERRORLEVEL% neq 0 (
	echo ビルドに失敗しました。
	exit /b 1
)

set NET48_EXE=.\bin\Release\net48\SpoCli.exe
set NET48_PUBLISH_DIR=.\bin\publish-net48
if not exist "%NET48_PUBLISH_DIR%" mkdir "%NET48_PUBLISH_DIR%"
copy /Y "%NET48_EXE%" "%NET48_PUBLISH_DIR%\SpoCli.exe" >nul

echo.
echo ビルド完了。
echo EXE: .\bin\publish-net48\SpoCli.exe
echo.
echo 使用方法: SpoCli.exe login --site https://contoso.sharepoint.com
goto end

:build_net10
echo Building net10 x64 self-contained single-file EXE...
dotnet publish SpoCli.csproj -c Release -f net10.0 -r win-x64 --self-contained=true -p:PublishSingleFile=true --output .\bin\publish
if %ERRORLEVEL% neq 0 (
	echo ビルドに失敗しました。
	exit /b 1
)

echo.
echo ビルド完了。
echo EXE: .\bin\publish\SpoCli.exe
echo.
echo 使用方法: SpoCli.exe login --site https://contoso.sharepoint.com

:end

endlocal
