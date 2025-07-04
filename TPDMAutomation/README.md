# TPDM Automation Application

A .NET Core console application that processes Excel files with delegate comments using ML.NET for intelligent classification into action categories.

## Overview

This application automatically processes Excel files containing employee/contractor data with "Delegate Comments" and classifies them into four action categories:
- **Add**: New employees/contractors being added
- **Update**: Existing records needing modifications
- **Term**: Employees/contractors being terminated
- **Other**: General inquiries or special cases

## Features

- ✅ **Multi-sheet Excel processing**: Handles Excel files with multiple worksheets
- ✅ **ML-based classification**: Uses ML.NET for intelligent comment analysis
- ✅ **Automatic fallback**: Defaults to "Add" when no delegate comments exist
- ✅ **Categorized output**: Creates separate Excel files for each action category
- ✅ **Comprehensive logging**: Detailed logging for monitoring and debugging
- ✅ **Robust error handling**: Graceful handling of various error scenarios

## Architecture

### Components

1. **MLService**: Handles ML.NET model training and prediction
2. **ExcelService**: Manages Excel file reading and writing using EPPlus
3. **ApplicationService**: Orchestrates the main processing workflow
4. **Models**: Data transfer objects for ML training and Excel processing

### Data Flow

```
Input Excel File → ML Model Training/Loading → Sheet Processing → 
Comment Classification → Categorized Output Files
```

## Prerequisites

- .NET 8.0 SDK or later
- Excel files in .xlsx format

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd TPDMAutomation
```

2. Restore dependencies:
```bash
dotnet restore
```

3. Build the application:
```bash
dotnet build
```

## Usage

### Command Line Arguments

```bash
dotnet run [inputFilePath] [outputDirectory]
```

- `inputFilePath`: Path to the Excel file to process
- `outputDirectory`: Directory where output files will be created (optional)

### Interactive Mode

If no arguments are provided, the application will prompt for input:

```bash
dotnet run
```

### Examples

1. **Process a file with default output location:**
```bash
dotnet run "/path/to/employee_data.xlsx"
```

2. **Process a file with custom output directory:**
```bash
dotnet run "/path/to/employee_data.xlsx" "/path/to/output"
```

3. **Interactive mode:**
```bash
dotnet run
# Follow the prompts to enter file paths
```

## Input File Requirements

### Excel File Structure

- **Format**: Excel 2007+ (.xlsx files)
- **Sheets**: Multiple sheets supported
- **Required Column**: "Delegate Comments" (case-insensitive)

### Sample Excel Structure

| Employee ID | Employee Name | Department | Delegate Comments | Manager |
|------------|---------------|------------|-------------------|---------|
| EMP001 | John Smith | IT | New employee starting next month | Jane Doe |
| EMP002 | Mary Johnson | HR | Employee information needs updating | Bob Wilson |
| EMP003 | David Brown | Finance | Employee is leaving the company | Alice Cooper |

## Output

The application creates four separate Excel files:

1. **`{InputFileName}_Add.xlsx`**: Records classified as "Add"
2. **`{InputFileName}_Update.xlsx`**: Records classified as "Update"
3. **`{InputFileName}_Term.xlsx`**: Records classified as "Term"
4. **`{InputFileName}_Other.xlsx`**: Records classified as "Other"

Each output file:
- Maintains the original sheet structure
- Includes all original columns
- Adds a "Predicted Action" column
- Preserves data formatting

## ML Model Training

### Training Data

The application uses training data from `Data/training_data.csv` with the following format:

```csv
Comment,Action
New employee starting next month,Add
Employee information needs updating,Update
Employee is leaving the company,Term
General inquiry about employee,Other
```

### Model Details

- **Algorithm**: SDCA Maximum Entropy (Multi-class Classification)
- **Features**: Text featurization of comment content
- **Output**: Action category with confidence scores
- **Model File**: Automatically saved as `Data/comment_classifier.zip`

### Retraining

The model automatically trains on first run if no saved model exists. To retrain:
1. Update the training data in `Data/training_data.csv`
2. Delete the existing model file `Data/comment_classifier.zip`
3. Run the application

## Configuration

### File Paths

- **Training Data**: `Data/training_data.csv`
- **Model File**: `Data/comment_classifier.zip`
- **Log Level**: Information (configurable in code)

### Customization

To modify classification categories or add new ones:
1. Update the training data CSV
2. Modify the `CreateCategorizedExcelFiles` method in `ExcelService`
3. Update the categories array in `ApplicationService`

## Error Handling

The application handles various error scenarios:

- **Missing input file**: Clear error message and validation
- **Invalid Excel format**: File extension and format validation
- **Missing delegate comments**: Automatic default to "Add" action
- **Empty worksheets**: Graceful handling with warnings
- **ML model issues**: Fallback to default classification
- **File access errors**: Comprehensive error logging

## Testing

### Creating Sample Data

Use the built-in sample data generator:

```csharp
SampleDataGenerator.CreateSampleExcelFile("sample_data.xlsx");
```

This creates a sample Excel file with:
- Multiple sheets
- Various comment types
- Missing delegate comments scenarios
- Different data structures

### Manual Testing

1. Create or use a sample Excel file
2. Run the application
3. Verify output files are created
4. Check logs for processing details
5. Validate categorization accuracy

## Logging

The application provides comprehensive logging:

- **Information**: Processing steps and summaries
- **Debug**: Detailed classification decisions
- **Warning**: Non-critical issues (e.g., empty comments)
- **Error**: Critical errors with stack traces

### Sample Log Output

```
[INFO] Application started.
[INFO] Reading Excel file: /path/to/input.xlsx
[INFO] Processing sheet: Employees with 10 rows
[INFO] Sheet 'Employees' processing complete:
[INFO]   Add: 4 rows
[INFO]   Update: 2 rows
[INFO]   Term: 2 rows
[INFO]   Other: 2 rows
[INFO] Created Add file: /path/to/output/input_Add.xlsx with 2 sheets.
```

## Performance Considerations

- **Memory Usage**: Large Excel files are processed efficiently using streaming
- **Processing Speed**: Typical rate of 1000+ rows per second
- **Model Size**: Lightweight model (~1MB) for fast loading
- **Scalability**: Can handle files with multiple sheets and thousands of rows

## Troubleshooting

### Common Issues

1. **"Training data file not found"**
   - Ensure `Data/training_data.csv` exists
   - Check file path and permissions

2. **"Model training failed"**
   - Verify training data format
   - Check for sufficient training examples
   - Review error logs for details

3. **"Cannot access output directory"**
   - Verify directory permissions
   - Ensure sufficient disk space
   - Check path validity

4. **Excel file reading errors**
   - Confirm file is not password-protected
   - Ensure file is not open in Excel
   - Verify file is not corrupted

### Support

For issues and questions:
1. Check the error logs for detailed information
2. Verify input file format and structure
3. Ensure all prerequisites are met
4. Review this documentation for guidance

## License

This project uses EPPlus for Excel processing under a non-commercial license. For commercial use, please ensure proper EPPlus licensing.

---

**Version**: 1.0  
**Last Updated**: 2024  
**Framework**: .NET 8.0