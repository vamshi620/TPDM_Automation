#!/bin/bash

# TPDM Automation Test Script
# This script runs a complete test of the TPDM Automation application

echo "========================================="
echo "TPDM Automation Application Test Suite"
echo "========================================="
echo

# Change to application directory
cd "$(dirname "$0")/.."

echo "Step 1: Building the application..."
dotnet build
if [ $? -ne 0 ]; then
    echo "❌ Build failed!"
    exit 1
fi
echo "✅ Build successful"
echo

echo "Step 2: Generating test data..."
dotnet run -- --generate-test-data
if [ $? -ne 0 ]; then
    echo "❌ Test data generation failed!"
    exit 1
fi
echo "✅ Test data generated successfully"
echo

echo "Step 3: Running application with test data..."
dotnet run -- -i "./Test/sample_input.xlsx" -o "./Test"
if [ $? -ne 0 ]; then
    echo "❌ Application execution failed!"
    exit 1
fi
echo "✅ Application executed successfully"
echo

echo "Step 4: Validating output files..."
expected_files=("Add_Records.xlsx" "Update_Records.xlsx" "Term_Records.xlsx" "Other_Records.xlsx" "ADD_Records.xlsx")
missing_files=()

for file in "${expected_files[@]}"; do
    if [ ! -f "./Test/$file" ]; then
        missing_files+=("$file")
    fi
done

if [ ${#missing_files[@]} -gt 0 ]; then
    echo "❌ Missing output files:"
    for file in "${missing_files[@]}"; do
        echo "   - $file"
    done
    exit 1
fi
echo "✅ All expected output files present"
echo

echo "Step 5: Test Results Summary"
echo "----------------------------"
echo "Input file: ./Test/sample_input.xlsx"
echo "Output directory: ./Test/"
echo
echo "Generated files:"
ls -la "./Test/"*.xlsx | while read -r line; do
    filename=$(echo "$line" | awk '{print $9}' | xargs basename)
    size=$(echo "$line" | awk '{print $5}')
    echo "  📄 $filename ($size bytes)"
done
echo

echo "========================================="
echo "🎉 All tests passed successfully!"
echo "========================================="
echo
echo "Test files are available in the ./Test/ directory"
echo "You can use these files to validate the application manually"
echo
echo "To clean up test files, run:"
echo "  rm -rf ./Test/"