using Microsoft.Extensions.Logging;
using TPDMAutomation.Models;
using TPDMAutomation.Services;

namespace TPDMAutomation.Services
{
    /// <summary>
    /// Main application service that orchestrates the entire processing workflow
    /// </summary>
    public class ApplicationService
    {
        private readonly MLService _mlService;
        private readonly ExcelService _excelService;
        private readonly ILogger<ApplicationService> _logger;

        public ApplicationService(
            MLService mlService, 
            ExcelService excelService, 
            ILogger<ApplicationService> logger)
        {
            _mlService = mlService;
            _excelService = excelService;
            _logger = logger;
        }

        /// <summary>
        /// Processes an Excel file according to the requirements:
        /// 1. Reads Excel file with multiple sheets
        /// 2. Processes delegate comments with ML predictions
        /// 3. Creates categorized output files
        /// </summary>
        /// <param name="config">Application configuration</param>
        /// <returns>True if processing was successful</returns>
        public async Task<bool> ProcessExcelFileAsync(AppConfig config)
        {
            try
            {
                _logger.LogInformation("Starting Excel file processing...");
                _logger.LogInformation($"Input file: {config.InputFilePath}");
                _logger.LogInformation($"Output directory: {config.OutputDirectory}");

                // Step 1: Ensure ML model is ready
                if (!await EnsureModelIsReadyAsync(config))
                {
                    _logger.LogError("ML model is not ready. Cannot proceed with processing.");
                    return false;
                }

                // Step 2: Read Excel file
                _logger.LogInformation("Reading Excel file...");
                var excelData = _excelService.ReadExcelFile(config.InputFilePath);
                
                if (!excelData.Any())
                {
                    _logger.LogError("No data found in Excel file or file could not be read.");
                    return false;
                }

                // Step 3: Process each sheet and predict actions
                _logger.LogInformation("Processing delegate comments and making predictions...");
                var processedData = ProcessDelegateComments(excelData);

                // Step 4: Create categorized output files
                _logger.LogInformation("Creating categorized output files...");
                var baseFileName = Path.GetFileNameWithoutExtension(config.InputFilePath);
                var success = _excelService.CreateCategorizedExcelFiles(
                    processedData, 
                    config.OutputDirectory, 
                    baseFileName);

                if (success)
                {
                    _logger.LogInformation("Excel file processing completed successfully.");
                    LogProcessingSummary(processedData);
                }
                else
                {
                    _logger.LogError("Failed to create output files.");
                }

                return success;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during Excel file processing.");
                return false;
            }
        }

        /// <summary>
        /// Ensures the ML model is trained and ready for predictions
        /// </summary>
        /// <param name="config">Application configuration</param>
        /// <returns>True if model is ready</returns>
        private async Task<bool> EnsureModelIsReadyAsync(AppConfig config)
        {
            try
            {
                // Check if model file exists
                if (File.Exists(config.ModelPath))
                {
                    _logger.LogInformation("Loading existing ML model...");
                    return _mlService.LoadModel(config.ModelPath);
                }

                // Model doesn't exist, need to train it
                _logger.LogInformation("ML model not found. Training new model...");
                
                if (!File.Exists(config.TrainingDataPath))
                {
                    _logger.LogError($"Training data file not found: {config.TrainingDataPath}");
                    return false;
                }

                // Train the model
                var trainingSuccess = await Task.Run(() => 
                    _mlService.TrainModel(config.TrainingDataPath, config.ModelPath));

                if (trainingSuccess)
                {
                    _logger.LogInformation("ML model trained and saved successfully.");
                }
                else
                {
                    _logger.LogError("Failed to train ML model.");
                }

                return trainingSuccess;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error ensuring ML model is ready.");
                return false;
            }
        }

        /// <summary>
        /// Processes delegate comments in all sheets and adds predictions
        /// </summary>
        /// <param name="excelData">Raw Excel data from all sheets</param>
        /// <returns>Processed data with predictions</returns>
        private Dictionary<string, List<ExcelRowData>> ProcessDelegateComments(
            Dictionary<string, List<ExcelRowData>> excelData)
        {
            var processedData = new Dictionary<string, List<ExcelRowData>>();

            foreach (var sheetData in excelData)
            {
                var sheetName = sheetData.Key;
                var rows = sheetData.Value;

                _logger.LogInformation($"Processing sheet: {sheetName} with {rows.Count} rows");

                var processedRows = new List<ExcelRowData>();

                foreach (var row in rows)
                {
                    var processedRow = ProcessSingleRow(row, sheetName);
                    processedRows.Add(processedRow);
                }

                processedData[sheetName] = processedRows;

                // Log sheet summary
                var categoryCounts = processedRows
                    .GroupBy(r => r.PredictedAction)
                    .ToDictionary(g => g.Key, g => g.Count());

                _logger.LogInformation($"Sheet '{sheetName}' processing complete:");
                foreach (var category in categoryCounts)
                {
                    _logger.LogInformation($"  {category.Key}: {category.Value} rows");
                }
            }

            return processedData;
        }

        /// <summary>
        /// Processes a single row and predicts the action
        /// </summary>
        /// <param name="row">Row data to process</param>
        /// <param name="sheetName">Name of the sheet (for logging)</param>
        /// <returns>Row with prediction added</returns>
        private ExcelRowData ProcessSingleRow(ExcelRowData row, string sheetName)
        {
            var processedRow = new ExcelRowData
            {
                RowNumber = row.RowNumber,
                DelegateComment = row.DelegateComment,
                OtherColumns = new Dictionary<string, object>(row.OtherColumns)
            };

            // Apply business rules for prediction
            if (string.IsNullOrWhiteSpace(row.DelegateComment))
            {
                // Rule: If no delegate comment, default to ADD
                processedRow.PredictedAction = "Add";
                _logger.LogDebug($"Sheet '{sheetName}', Row {row.RowNumber}: No delegate comment, defaulting to Add");
            }
            else
            {
                // Use ML model to predict
                processedRow.PredictedAction = _mlService.PredictAction(row.DelegateComment);
                _logger.LogDebug($"Sheet '{sheetName}', Row {row.RowNumber}: '{row.DelegateComment}' -> {processedRow.PredictedAction}");
            }

            return processedRow;
        }

        /// <summary>
        /// Logs a summary of the processing results
        /// </summary>
        /// <param name="processedData">Processed data to summarize</param>
        private void LogProcessingSummary(Dictionary<string, List<ExcelRowData>> processedData)
        {
            _logger.LogInformation("=== PROCESSING SUMMARY ===");
            
            var totalRows = processedData.Values.Sum(rows => rows.Count);
            _logger.LogInformation($"Total rows processed: {totalRows}");
            _logger.LogInformation($"Total sheets processed: {processedData.Count}");

            // Overall category distribution
            var allRows = processedData.Values.SelectMany(rows => rows);
            var overallCategoryCounts = allRows
                .GroupBy(r => r.PredictedAction)
                .ToDictionary(g => g.Key, g => g.Count());

            _logger.LogInformation("Overall category distribution:");
            foreach (var category in overallCategoryCounts.OrderByDescending(c => c.Value))
            {
                var percentage = (double)category.Value / totalRows * 100;
                _logger.LogInformation($"  {category.Key}: {category.Value} rows ({percentage:F1}%)");
            }

            // Per-sheet breakdown
            _logger.LogInformation("Per-sheet breakdown:");
            foreach (var sheetData in processedData)
            {
                var sheetCategoryCounts = sheetData.Value
                    .GroupBy(r => r.PredictedAction)
                    .ToDictionary(g => g.Key, g => g.Count());

                _logger.LogInformation($"  {sheetData.Key}: {sheetData.Value.Count} rows");
                foreach (var category in sheetCategoryCounts)
                {
                    _logger.LogInformation($"    {category.Key}: {category.Value}");
                }
            }

            _logger.LogInformation("=== END SUMMARY ===");
        }
    }
}