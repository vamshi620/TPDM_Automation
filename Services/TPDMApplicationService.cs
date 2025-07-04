using Microsoft.Extensions.Logging;
using TPDMAutomation.Models;

namespace TPDMAutomation.Services
{
    /// <summary>
    /// Main application service that orchestrates the entire TPDM automation process
    /// </summary>
    public class TPDMApplicationService
    {
        private readonly MLModelService _mlModelService;
        private readonly ExcelProcessingService _excelProcessingService;
        private readonly ILogger<TPDMApplicationService> _logger;

        /// <summary>
        /// Initializes a new instance of the TPDMApplicationService
        /// </summary>
        /// <param name="mlModelService">ML model service</param>
        /// <param name="excelProcessingService">Excel processing service</param>
        /// <param name="logger">Logger instance</param>
        public TPDMApplicationService(
            MLModelService mlModelService,
            ExcelProcessingService excelProcessingService,
            ILogger<TPDMApplicationService> logger)
        {
            _mlModelService = mlModelService;
            _excelProcessingService = excelProcessingService;
            _logger = logger;
        }

        /// <summary>
        /// Processes Excel file according to TPDM automation requirements
        /// </summary>
        /// <param name="config">Processing configuration</param>
        /// <returns>True if processing was successful, false otherwise</returns>
        public async Task<bool> ProcessExcelFileAsync(ProcessingConfig config)
        {
            try
            {
                _logger.LogInformation("Starting TPDM automation process");
                _logger.LogInformation("Input file: {InputPath}", config.InputExcelPath);
                _logger.LogInformation("Output directory: {OutputDirectory}", config.OutputDirectory);

                // Step 1: Validate inputs
                if (!ValidateInputs(config))
                {
                    return false;
                }

                // Step 2: Train ML model if not already trained
                if (!_mlModelService.IsModelTrained())
                {
                    _logger.LogInformation("ML model not trained. Training model...");
                    var trainingDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "training_data.csv");
                    
                    if (!_mlModelService.TrainModel(trainingDataPath))
                    {
                        _logger.LogError("Failed to train ML model");
                        return false;
                    }
                }

                // Step 3: Read Excel file and extract data
                _logger.LogInformation("Reading and processing Excel file...");
                var excelData = _excelProcessingService.ReadExcelFile(config);

                if (!excelData.Any())
                {
                    _logger.LogWarning("No data found in Excel file");
                    return false;
                }

                // Step 4: Process delegate comments and make predictions
                _logger.LogInformation("Processing delegate comments and making predictions...");
                await ProcessDelegateCommentsAsync(excelData, config);

                // Step 5: Create output files grouped by categories
                _logger.LogInformation("Creating output Excel files...");
                _excelProcessingService.CreateOutputFiles(excelData, config);

                // Step 6: Log summary
                LogProcessingSummary(excelData);

                _logger.LogInformation("TPDM automation process completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred during TPDM automation process");
                return false;
            }
        }

        /// <summary>
        /// Validates input configuration
        /// </summary>
        /// <param name="config">Processing configuration</param>
        /// <returns>True if inputs are valid, false otherwise</returns>
        private bool ValidateInputs(ProcessingConfig config)
        {
            if (string.IsNullOrWhiteSpace(config.InputExcelPath))
            {
                _logger.LogError("Input Excel path is required");
                return false;
            }

            if (!File.Exists(config.InputExcelPath))
            {
                _logger.LogError("Input Excel file does not exist: {FilePath}", config.InputExcelPath);
                return false;
            }

            if (string.IsNullOrWhiteSpace(config.OutputDirectory))
            {
                _logger.LogError("Output directory is required");
                return false;
            }

            // Create output directory if it doesn't exist
            try
            {
                Directory.CreateDirectory(config.OutputDirectory);
                _logger.LogInformation("Output directory created/verified: {OutputDirectory}", config.OutputDirectory);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to create output directory: {OutputDirectory}", config.OutputDirectory);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Processes delegate comments and makes category predictions
        /// </summary>
        /// <param name="excelData">Excel data to process</param>
        /// <param name="config">Processing configuration</param>
        private async Task ProcessDelegateCommentsAsync(List<ExcelRowData> excelData, ProcessingConfig config)
        {
            int processedCount = 0;
            int totalCount = excelData.Count;

            foreach (var rowData in excelData)
            {
                // If we already have a predicted category (e.g., default for missing column), skip prediction
                if (!string.IsNullOrEmpty(rowData.PredictedCategory))
                {
                    processedCount++;
                    continue;
                }

                // If delegate comment is empty or null, use default category
                if (string.IsNullOrWhiteSpace(rowData.DelegateComment))
                {
                    rowData.PredictedCategory = config.DefaultCategory;
                    _logger.LogDebug("Using default category {DefaultCategory} for empty comment in sheet {SheetName}, row {RowIndex}", 
                        config.DefaultCategory, rowData.SheetName, rowData.RowIndex);
                }
                else
                {
                    // Use ML model to predict category
                    rowData.PredictedCategory = _mlModelService.PredictCategory(rowData.DelegateComment);
                    _logger.LogDebug("Predicted category {Category} for comment '{Comment}' in sheet {SheetName}, row {RowIndex}", 
                        rowData.PredictedCategory, rowData.DelegateComment, rowData.SheetName, rowData.RowIndex);
                }

                processedCount++;

                // Log progress every 100 rows
                if (processedCount % 100 == 0)
                {
                    _logger.LogInformation("Processed {ProcessedCount}/{TotalCount} rows", processedCount, totalCount);
                }

                // Add small delay to prevent overwhelming the system
                if (processedCount % 50 == 0)
                {
                    await Task.Delay(10);
                }
            }

            _logger.LogInformation("Completed processing all {TotalCount} rows", totalCount);
        }

        /// <summary>
        /// Logs a summary of the processing results
        /// </summary>
        /// <param name="excelData">Processed Excel data</param>
        private void LogProcessingSummary(List<ExcelRowData> excelData)
        {
            var categoryCounts = excelData.GroupBy(d => d.PredictedCategory)
                .ToDictionary(g => g.Key, g => g.Count());

            _logger.LogInformation("Processing Summary:");
            _logger.LogInformation("Total rows processed: {TotalRows}", excelData.Count);
            
            foreach (var categoryCount in categoryCounts.OrderBy(kv => kv.Key))
            {
                _logger.LogInformation("  {Category}: {Count} records", categoryCount.Key, categoryCount.Value);
            }

            var sheetCounts = excelData.GroupBy(d => d.SheetName)
                .ToDictionary(g => g.Key, g => g.Count());

            _logger.LogInformation("Records by sheet:");
            foreach (var sheetCount in sheetCounts.OrderBy(kv => kv.Key))
            {
                _logger.LogInformation("  {SheetName}: {Count} records", sheetCount.Key, sheetCount.Value);
            }
        }
    }
}