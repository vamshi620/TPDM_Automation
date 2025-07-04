# TPDM Automation Test Script for Windows
# This script runs a complete test of the TPDM Automation application

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "TPDM Automation Application Test Suite" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host

# Change to application directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location (Join-Path $scriptPath "..")

Write-Host "Step 1: Building the application..." -ForegroundColor Yellow
$buildResult = & dotnet build 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Build failed!" -ForegroundColor Red
    Write-Host $buildResult -ForegroundColor Red
    exit 1
}
Write-Host "✅ Build successful" -ForegroundColor Green
Write-Host

Write-Host "Step 2: Generating test data..." -ForegroundColor Yellow
$generateResult = & dotnet run -- --generate-test-data 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Test data generation failed!" -ForegroundColor Red
    Write-Host $generateResult -ForegroundColor Red
    exit 1
}
Write-Host "✅ Test data generated successfully" -ForegroundColor Green
Write-Host

Write-Host "Step 3: Running application with test data..." -ForegroundColor Yellow
$runResult = & dotnet run -- -i "./Test/sample_input.xlsx" -o "./Test" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Application execution failed!" -ForegroundColor Red
    Write-Host $runResult -ForegroundColor Red
    exit 1
}
Write-Host "✅ Application executed successfully" -ForegroundColor Green
Write-Host

Write-Host "Step 4: Validating output files..." -ForegroundColor Yellow
$expectedFiles = @("Add_Records.xlsx", "Update_Records.xlsx", "Term_Records.xlsx", "Other_Records.xlsx", "ADD_Records.xlsx")
$missingFiles = @()

foreach ($file in $expectedFiles) {
    $filePath = Join-Path "./Test" $file
    if (!(Test-Path $filePath)) {
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "❌ Missing output files:" -ForegroundColor Red
    foreach ($file in $missingFiles) {
        Write-Host "   - $file" -ForegroundColor Red
    }
    exit 1
}
Write-Host "✅ All expected output files present" -ForegroundColor Green
Write-Host

Write-Host "Step 5: Test Results Summary" -ForegroundColor Yellow
Write-Host "----------------------------" -ForegroundColor Yellow
Write-Host "Input file: ./Test/sample_input.xlsx"
Write-Host "Output directory: ./Test/"
Write-Host

Write-Host "Generated files:"
Get-ChildItem "./Test/*.xlsx" | ForEach-Object {
    $size = [math]::Round($_.Length / 1KB, 2)
    Write-Host "  📄 $($_.Name) ($size KB)" -ForegroundColor Cyan
}
Write-Host

Write-Host "=========================================" -ForegroundColor Green
Write-Host "🎉 All tests passed successfully!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host

Write-Host "Test files are available in the ./Test/ directory"
Write-Host "You can use these files to validate the application manually"
Write-Host
Write-Host "To clean up test files, run:"
Write-Host "  Remove-Item -Recurse -Force ./Test/" -ForegroundColor Yellow