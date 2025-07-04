using OfficeOpenXml;
using Microsoft.Extensions.Logging;
using TPDMAutomation.Models;

namespace TPDMAutomation.Services
{
    /// <summary>
    /// Service for reading and writing Excel files
    /// </summary>
    public class ExcelService
    {
        private readonly ILogger<ExcelService> _logger;
        private const string DelegateCommentsColumn = "Delegate Comments";

        public ExcelService(ILogger<ExcelService> logger)
        {
            _logger = logger;
            // Set the license context for EPPlus (required for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Reads all sheets from an Excel file and extracts data with delegate comments
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <returns>Dictionary with sheet names as keys and list of row data as values</returns>
        public Dictionary<string, List<ExcelRowData>> ReadExcelFile(string filePath)
        {
            var result = new Dictionary<string, List<ExcelRowData>>();

            try
            {
                if (!File.Exists(filePath))
                {
                    _logger.LogError($"Excel file not found at: {filePath}");
                    return result;
                }

                _logger.LogInformation($"Reading Excel file: {filePath}");

                using var package = new ExcelPackage(new FileInfo(filePath));
                
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    _logger.LogInformation($"Processing sheet: {worksheet.Name}");
                    
                    var sheetData = ReadWorksheet(worksheet);
                    result[worksheet.Name] = sheetData;
                    
                    _logger.LogInformation($"Sheet '{worksheet.Name}' processed. Rows: {sheetData.Count}");
                }

                _logger.LogInformation($"Excel file processed successfully. Total sheets: {result.Count}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error reading Excel file: {filePath}");
            }

            return result;
        }

        /// <summary>
        /// Reads data from a single worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet to read</param>
        /// <returns>List of row data from the worksheet</returns>
        private List<ExcelRowData> ReadWorksheet(ExcelWorksheet worksheet)
        {
            var data = new List<ExcelRowData>();

            try
            {
                if (worksheet.Dimension == null)
                {
                    _logger.LogWarning($"Worksheet '{worksheet.Name}' is empty.");
                    return data;
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                if (rowCount < 2) // Need at least header and one data row
                {
                    _logger.LogWarning($"Worksheet '{worksheet.Name}' has insufficient data (less than 2 rows).");
                    return data;
                }

                // Find column headers in the first row
                var headers = new Dictionary<int, string>();
                int delegateCommentsColumn = -1;

                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim() ?? "";
                    headers[col] = headerValue;
                    
                    if (string.Equals(headerValue, DelegateCommentsColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        delegateCommentsColumn = col;
                    }
                }

                _logger.LogInformation($"Found {headers.Count} columns. Delegate Comments column: {(delegateCommentsColumn > 0 ? delegateCommentsColumn.ToString() : "Not Found")}");

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new ExcelRowData
                    {
                        RowNumber = row,
                        DelegateComment = delegateCommentsColumn > 0 
                            ? worksheet.Cells[row, delegateCommentsColumn].Value?.ToString()?.Trim() ?? ""
                            : "",
                        OtherColumns = new Dictionary<string, object>()
                    };

                    // Read all other columns
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (col != delegateCommentsColumn)
                        {
                            var columnName = headers.ContainsKey(col) ? headers[col] : $"Column{col}";
                            var cellValue = worksheet.Cells[row, col].Value;
                            rowData.OtherColumns[columnName] = cellValue ?? "";
                        }
                    }

                    data.Add(rowData);
                }

                _logger.LogInformation($"Read {data.Count} data rows from worksheet '{worksheet.Name}'.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error reading worksheet '{worksheet.Name}'.");
            }

            return data;
        }

        /// <summary>
        /// Creates separate Excel files for each action category
        /// </summary>
        /// <param name="processedData">Dictionary with sheet names and processed row data</param>
        /// <param name="outputDirectory">Directory where to save the output files</param>
        /// <param name="baseFileName">Base name for output files</param>
        /// <returns>True if files were created successfully</returns>
        public bool CreateCategorizedExcelFiles(
            Dictionary<string, List<ExcelRowData>> processedData, 
            string outputDirectory, 
            string baseFileName)
        {
            try
            {
                _logger.LogInformation("Creating categorized Excel files...");

                // Ensure output directory exists
                Directory.CreateDirectory(outputDirectory);

                var categories = new[] { "Add", "Update", "Term", "Other" };

                foreach (var category in categories)
                {
                    var fileName = $"{baseFileName}_{category}.xlsx";
                    var filePath = Path.Combine(outputDirectory, fileName);

                    CreateCategoryFile(processedData, category, filePath);
                }

                _logger.LogInformation("All categorized Excel files created successfully.");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating categorized Excel files.");
                return false;
            }
        }

        /// <summary>
        /// Creates an Excel file for a specific category
        /// </summary>
        /// <param name="allData">All processed data</param>
        /// <param name="category">Category to filter by</param>
        /// <param name="filePath">Output file path</param>
        private void CreateCategoryFile(
            Dictionary<string, List<ExcelRowData>> allData, 
            string category, 
            string filePath)
        {
            try
            {
                using var package = new ExcelPackage();

                foreach (var sheetData in allData)
                {
                    var sheetName = sheetData.Key;
                    var rows = sheetData.Value.Where(r => r.PredictedAction.Equals(category, StringComparison.OrdinalIgnoreCase)).ToList();

                    if (rows.Any())
                    {
                        var worksheet = package.Workbook.Worksheets.Add(sheetName);
                        WriteWorksheetData(worksheet, rows, category);
                    }
                }

                if (package.Workbook.Worksheets.Count > 0)
                {
                    package.SaveAs(new FileInfo(filePath));
                    _logger.LogInformation($"Created {category} file: {filePath} with {package.Workbook.Worksheets.Count} sheets.");
                }
                else
                {
                    _logger.LogInformation($"No data found for category {category}. File not created.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error creating category file for {category}.");
            }
        }

        /// <summary>
        /// Writes data to a worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="rows">Data rows to write</param>
        /// <param name="category">Category being written</param>
        private void WriteWorksheetData(ExcelWorksheet worksheet, List<ExcelRowData> rows, string category)
        {
            if (!rows.Any()) return;

            // Get all column names from the first row
            var allColumns = rows.First().OtherColumns.Keys.ToList();
            allColumns.Add(DelegateCommentsColumn);
            allColumns.Add("Predicted Action");

            // Write headers
            for (int i = 0; i < allColumns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = allColumns[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            // Write data rows
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                int excelRow = rowIndex + 2; // +2 because Excel is 1-based and we have a header row

                for (int colIndex = 0; colIndex < allColumns.Count; colIndex++)
                {
                    var columnName = allColumns[colIndex];
                    object value = "";

                    if (columnName == DelegateCommentsColumn)
                    {
                        value = row.DelegateComment;
                    }
                    else if (columnName == "Predicted Action")
                    {
                        value = row.PredictedAction;
                    }
                    else if (row.OtherColumns.ContainsKey(columnName))
                    {
                        value = row.OtherColumns[columnName];
                    }

                    worksheet.Cells[excelRow, colIndex + 1].Value = value;
                }
            }

            // Auto-fit columns
            worksheet.Cells.AutoFitColumns();
        }
    }
}