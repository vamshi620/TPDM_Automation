using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using TPDMAutomation.Models;

namespace TPDMAutomation.Services
{
    /// <summary>
    /// Service for processing Excel files and extracting/writing data
    /// </summary>
    public class ExcelProcessingService
    {
        private readonly ILogger<ExcelProcessingService> _logger;

        /// <summary>
        /// Initializes a new instance of the ExcelProcessingService
        /// </summary>
        /// <param name="logger">Logger instance</param>
        public ExcelProcessingService(ILogger<ExcelProcessingService> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Reads Excel file and extracts data from all sheets
        /// </summary>
        /// <param name="config">Processing configuration</param>
        /// <returns>List of Excel row data</returns>
        public List<ExcelRowData> ReadExcelFile(ProcessingConfig config)
        {
            var result = new List<ExcelRowData>();

            try
            {
                _logger.LogInformation("Reading Excel file: {FilePath}", config.InputExcelPath);

                using var workbook = new XLWorkbook(config.InputExcelPath);

                foreach (var worksheet in workbook.Worksheets)
                {
                    _logger.LogInformation("Processing sheet: {SheetName}", worksheet.Name);

                    var sheetData = ProcessWorksheet(worksheet, config);
                    result.AddRange(sheetData);
                }

                _logger.LogInformation("Successfully read {RowCount} rows from {SheetCount} sheets", 
                    result.Count, workbook.Worksheets.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading Excel file: {FilePath}", config.InputExcelPath);
                throw;
            }

            return result;
        }

        /// <summary>
        /// Processes a single worksheet and extracts data
        /// </summary>
        /// <param name="worksheet">The worksheet to process</param>
        /// <param name="config">Processing configuration</param>
        /// <returns>List of Excel row data from the sheet</returns>
        private List<ExcelRowData> ProcessWorksheet(IXLWorksheet worksheet, ProcessingConfig config)
        {
            var result = new List<ExcelRowData>();

            try
            {
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    _logger.LogWarning("Sheet {SheetName} is empty", worksheet.Name);
                    return result;
                }

                // Get header row
                var headerRow = usedRange.FirstRow();
                var headers = headerRow.Cells().Select(c => c.GetString()).ToList();

                _logger.LogDebug("Found headers in sheet {SheetName}: {Headers}", worksheet.Name, string.Join(", ", headers));

                // Find delegate comments column index
                var delegateCommentsIndex = FindColumnIndex(headers, config.DelegateCommentsColumnName);
                bool hasDelegateComments = delegateCommentsIndex >= 0;

                _logger.LogInformation("Sheet {SheetName} - Delegate Comments column found: {Found} at index: {Index}", 
                    worksheet.Name, hasDelegateComments, delegateCommentsIndex);

                // Process data rows
                for (int rowIndex = 2; rowIndex <= usedRange.RowCount(); rowIndex++)
                {
                    var row = worksheet.Row(rowIndex);
                    var rowData = new ExcelRowData
                    {
                        RowIndex = rowIndex,
                        SheetName = worksheet.Name,
                        ColumnValues = new Dictionary<string, object>()
                    };

                    // Extract all column values
                    for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                    {
                        var cellValue = row.Cell(colIndex + 1).Value;
                        rowData.ColumnValues[headers[colIndex]] = cellValue;
                    }

                    // Extract delegate comment
                    if (hasDelegateComments && delegateCommentsIndex < headers.Count)
                    {
                        rowData.DelegateComment = row.Cell(delegateCommentsIndex + 1).GetString();
                    }
                    else
                    {
                        // No delegate comments column found, use default category
                        rowData.DelegateComment = "";
                        rowData.PredictedCategory = config.DefaultCategory;
                    }

                    result.Add(rowData);
                }

                _logger.LogInformation("Processed {RowCount} data rows from sheet {SheetName}", result.Count, worksheet.Name);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing worksheet: {SheetName}", worksheet.Name);
                throw;
            }

            return result;
        }

        /// <summary>
        /// Finds the index of a column by name (case-insensitive)
        /// </summary>
        /// <param name="headers">List of header names</param>
        /// <param name="columnName">Column name to find</param>
        /// <returns>Column index or -1 if not found</returns>
        private int FindColumnIndex(List<string> headers, string columnName)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (string.Equals(headers[i], columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Creates output Excel files grouped by predicted categories
        /// </summary>
        /// <param name="data">Processed Excel data with predictions</param>
        /// <param name="config">Processing configuration</param>
        public void CreateOutputFiles(List<ExcelRowData> data, ProcessingConfig config)
        {
            try
            {
                _logger.LogInformation("Creating output Excel files in directory: {OutputDirectory}", config.OutputDirectory);

                // Group data by predicted category
                var groupedData = data.GroupBy(d => d.PredictedCategory).ToList();

                foreach (var group in groupedData)
                {
                    var category = group.Key;
                    var categoryData = group.ToList();

                    var outputFileName = Path.Combine(config.OutputDirectory, $"{category}_Records.xlsx");
                    CreateCategoryExcelFile(categoryData, outputFileName, category);

                    _logger.LogInformation("Created output file: {FileName} with {RecordCount} records", 
                        outputFileName, categoryData.Count);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating output files");
                throw;
            }
        }

        /// <summary>
        /// Creates an Excel file for a specific category
        /// </summary>
        /// <param name="categoryData">Data for the category</param>
        /// <param name="outputFileName">Output file name</param>
        /// <param name="category">Category name</param>
        private void CreateCategoryExcelFile(List<ExcelRowData> categoryData, string outputFileName, string category)
        {
            using var workbook = new XLWorkbook();

            // Group by sheet name
            var sheetGroups = categoryData.GroupBy(d => d.SheetName);

            foreach (var sheetGroup in sheetGroups)
            {
                var sheetName = sheetGroup.Key;
                var sheetData = sheetGroup.ToList();

                // Create worksheet
                var worksheet = workbook.Worksheets.Add(sheetName);

                if (sheetData.Any())
                {
                    // Create headers
                    var headers = sheetData.First().ColumnValues.Keys.ToList();
                    headers.Add("Predicted Category"); // Add the prediction column

                    // Write headers
                    for (int i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    }

                    // Write data rows
                    for (int rowIndex = 0; rowIndex < sheetData.Count; rowIndex++)
                    {
                        var rowData = sheetData[rowIndex];
                        
                        // Write original column values
                        for (int colIndex = 0; colIndex < headers.Count - 1; colIndex++)
                        {
                            var columnName = headers[colIndex];
                            if (rowData.ColumnValues.ContainsKey(columnName))
                            {
                                worksheet.Cell(rowIndex + 2, colIndex + 1).Value = rowData.ColumnValues[columnName].ToString();
                            }
                        }

                        // Write predicted category
                        worksheet.Cell(rowIndex + 2, headers.Count).Value = rowData.PredictedCategory;
                    }

                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();
                }
            }

            // Save the workbook
            workbook.SaveAs(outputFileName);
        }
    }
}