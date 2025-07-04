# TPDM Automation Application

## Overview

The TPDM Automation application is a .NET Core console application that processes Excel files containing delegate comments and uses Machine Learning (ML.NET) to classify them into categories: **Add**, **Update**, **Term**, and **Other**.

## Features

- **ML.NET Integration**: Trains a text classification model to predict categories based on delegate comments
- **Multi-Sheet Excel Processing**: Reads Excel files with multiple worksheets
- **Intelligent Column Detection**: Automatically finds "Delegate Comments" columns in each sheet
- **Category-Based Output**: Generates separate Excel files for each predicted category
- **Flexible Configuration**: Command-line arguments and interactive prompts for configuration
- **Comprehensive Logging**: Detailed logging for monitoring and debugging
- **Error Handling**: Robust error handling with meaningful error messages

## Architecture

### Core Components

1. **Models** (`Models/`)
   - `CommentData.cs`: ML.NET training data model
   - `ExcelModels.cs`: Excel processing data models

2. **Services** (`Services/`)
   - `MLModelService.cs`: ML.NET model training and prediction
   - `ExcelProcessingService.cs`: Excel file reading and writing
   - `TPDMApplicationService.cs`: Main orchestration service

3. **Utilities** (`Utils/`)
   - `ConfigurationHelper.cs`: Configuration parsing and validation

4. **Data** (`Data/`)
   - `training_data.csv`: Sample training data for the ML model

## Requirements

- .NET 8.0 or later
- Excel files in `.xlsx` format
- Windows, macOS, or Linux operating system

## Installation

1. Clone the repository
2. Navigate to the project directory
3. Run `dotnet restore` to restore dependencies
4. Run `dotnet build` to build the application

## Usage

### Command Line Arguments

```bash
TPDMAutomation --input <excel_file> --output <output_directory> [options]
```

#### Required Arguments:
- `--input, -i`: Path to the input Excel file
- `--output, -o`: Output directory for generated Excel files

#### Optional Arguments:
- `--column, -c`: Name of the delegate comments column (default: "Delegate Comments")
- `--default, -d`: Default category when no comments found (default: "ADD")
- `--help, -h`: Show help message

### Examples

```bash
# Basic usage
TPDMAutomation -i "C:\data\input.xlsx" -o "C:\output"

# With custom column name and default category
TPDMAutomation --input "./data.xlsx" --output "./results" --column "Comments" --default "UPDATE"

# Interactive mode (prompts for missing parameters)
TPDMAutomation
```

## Processing Logic

1. **Model Training**: The application first trains an ML.NET text classification model using the sample training data
2. **Excel Reading**: Reads all worksheets from the input Excel file
3. **Column Detection**: Searches for the "Delegate Comments" column in each sheet
4. **Comment Processing**: 
   - If "Delegate Comments" column exists: Uses ML model to predict category
   - If column is missing: Assigns the default category ("ADD") to all rows
5. **Output Generation**: Creates four separate Excel files based on predicted categories:
   - `Add_Records.xlsx`
   - `Update_Records.xlsx` 
   - `Term_Records.xlsx`
   - `Other_Records.xlsx`

## ML Model Categories

The ML model classifies comments into four categories:

- **Add**: New employee additions, registrations, onboarding
- **Update**: Information changes, modifications, revisions
- **Term**: Terminations, departures, contract endings
- **Other**: Reviews, investigations, pending items

## Configuration

### ProcessingConfig Properties

- `InputExcelPath`: Path to the input Excel file
- `OutputDirectory`: Directory where output files will be created
- `DelegateCommentsColumnName`: Name of the column containing comments (default: "Delegate Comments")
- `DefaultCategory`: Category to use when no comments column is found (default: "ADD")

## Error Handling

The application includes comprehensive error handling for:

- Missing or invalid input files
- Invalid Excel file formats
- Missing output directory permissions
- ML model training failures
- File I/O errors

## Logging

The application uses Microsoft.Extensions.Logging with console output. Log levels include:

- **Information**: General processing status
- **Debug**: Detailed processing information
- **Warning**: Non-critical issues
- **Error**: Critical errors and exceptions

## Extending the Application

### Adding New Categories

1. Update the training data in `Data/training_data.csv`
2. Add sample comments for the new category
3. The ML model will automatically learn the new category

### Custom Column Processing

Modify the `ExcelProcessingService` to handle additional column types or processing logic.

### Different ML Models

Replace the `MLModelService` implementation to use different ML.NET trainers or external ML services.

## Performance Considerations

- The application processes rows in batches with small delays to prevent system overload
- Memory usage is optimized for large Excel files
- Progress logging every 100 rows for long-running operations

## Troubleshooting

### Common Issues

1. **"Training data not found"**: Ensure `Data/training_data.csv` exists in the application directory
2. **"Excel file cannot be opened"**: Verify the file is not open in Excel and has proper permissions
3. **"Output directory access denied"**: Ensure the application has write permissions to the output directory

### Debug Mode

Run with verbose logging by modifying the log level in `Program.cs`:

```csharp
builder.SetMinimumLevel(LogLevel.Debug);
```

## License

This project is part of the TPDM Automation system. Please refer to the project license for usage terms.