# TPDM_Automation

A .NET Core console application that uses ML.NET to classify Excel data and automate document processing.

## Features

1. **ML.NET Model**: Trains with comment data to predict categories (Add, Update, Term, Other)
2. **Multi-sheet Excel Processing**: Handles Excel files with multiple worksheets
3. **Delegate Comments Classification**: Analyzes "Delegate Comments" column for predictions
4. **Default Category Assignment**: Assigns "ADD" to sheets without delegate comments
5. **Automated Output Generation**: Creates separate Excel files for each category
6. **Comprehensive Testing**: Includes sample data generation and validation tools

## Requirements

You are .net core experienced developer agent.need to develop a .net core console application with below requirments.

1. Create a ML.Net model which will be trained with data from excel, with comments values and based on comments it need to predict Add,Update,Term and other.
2. My application need to take a excel from share path.
3. Excel from share path will have multiple sheets.
4. Each sheet will have a column called "Delegate Comments". 
5. My application need to read through "delegate comments" from each row and predict the values Like add,Update,term and other, then create a new column in each sheet with predicated values.
6. If excel is not having any "Delegate Comments" then consider it as ADD for all rows in a sheet.
7. after predicting values need to create a 4 excels as output in same share path. each excel will consits of correspondent predicted row values. like add records in one excel, update as one excel, term as one excel and other as one excel.
8. generate application with proper documentation and comments

## Quick Start

### Generate Sample Test Data
```bash
dotnet run -- --generate-test-data
```

### Run with Sample Data
```bash
dotnet run -- -i "./Test/sample_input.xlsx" -o "./Test"
```

### Run Automated Test Suite
```bash
# Linux/macOS
./Test/run_test.sh

# Windows
.\Test\run_test.ps1
```

## Testing and Validation

The application includes comprehensive testing capabilities:

- **Sample Data Generation**: Creates realistic test Excel files with multiple sheets
- **Automated Test Scripts**: Bash and PowerShell scripts for complete testing
- **Validation Reports**: Detailed test results and validation documentation
- **Test Coverage**: All application features tested with sample data

See `Test/README.md` and `Test/VALIDATION_REPORT.md` for detailed testing information.
