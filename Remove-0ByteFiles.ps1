# Define the target directory
$targetPath = "path"

# Uncomment the lines below to actually delete (after you've reviewed the test results)
# Write-Host "`n=== Deleting 0-byte files ===" -ForegroundColor Green
 Get-ChildItem -Path $targetPath -File -Recurse | Where-Object { $_.Length -eq 0 } | Remove-Item -Force

# Write-Host "`n=== Deleting empty folders ===" -ForegroundColor Green
 Get-ChildItem -Path $targetPath -Directory -Recurse | Where-Object { @(Get-ChildItem -Path $_.FullName -Recurse).Count -eq 0 } | Remove-Item -Force
