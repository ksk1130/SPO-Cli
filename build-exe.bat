@echo off
REM x64用のシングル実行可能ファイル（EXE）をビルド
REM .NET Runtimeなしで動作する自給自足型EXE

setlocal enabledelayexpansion

echo Building x64 self-contained EXE...
dotnet publish -c Release -r win-x64 --self-contained=true -p:PublishSingleFile=true --output ./bin/publish

if %ERRORLEVEL% equ 0 (
	echo.
	echo ビルド完了。
	echo EXE: .\bin\publish\SpoCli.exe
	echo.
	echo 使用方法: SpoCli.exe login --site https://contoso.sharepoint.com
) else (
	echo ビルドに失敗しました。
	exit /b 1
)

endlocal
