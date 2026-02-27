# テストデータ作成スクリプト
# dirA/dirB/dirC のようなフォルダ構造を作成

$testDataRoot = "testdata"

# クリーンアップ
if (Test-Path $testDataRoot) {
    Remove-Item -Recurse -Force $testDataRoot
}

# フォルダ構造作成
New-Item -ItemType Directory -Path "$testDataRoot/dirA" -Force | Out-Null
New-Item -ItemType Directory -Path "$testDataRoot/dirA/dirB" -Force | Out-Null
New-Item -ItemType Directory -Path "$testDataRoot/dirA/dirB/dirC" -Force | Out-Null

# テストファイル作成
Set-Content -Path "$testDataRoot/dirA/a.txt" -Value "This is file a.txt in dirA"
Set-Content -Path "$testDataRoot/dirA/dirB/b.txt" -Value "This is file b.txt in dirB"
Set-Content -Path "$testDataRoot/dirA/dirB/dirC/c.txt" -Value "This is file c.txt in dirC"

Write-Host "Test data structure created:" -ForegroundColor Green
Write-Host ""
Get-ChildItem -Path $testDataRoot -Recurse | ForEach-Object {
    $indent = "  " * ($_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count - $testDataRoot.Split([IO.Path]::DirectorySeparatorChar).Count - 1)
    if ($_.PSIsContainer) {
        Write-Host "$indent[DIR]  $($_.Name)" -ForegroundColor Cyan
    } else {
        Write-Host "$indent[FILE] $($_.Name)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Green
Write-Host "1. Start the mock server: dotnet run --project MockSpoServer.csproj"
Write-Host "2. Run SpoCli against mock server:"
Write-Host "   spocli --recursive cp http://localhost:5000/sites/testsite/Shared%20Documents/dirA/ ./output"
