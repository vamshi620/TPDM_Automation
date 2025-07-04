using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using TPDMAutomation.Models;
using TPDMAutomation.Services;

namespace TPDMAutomation
{
    /// <summary>
    /// TPDM Automation Console Application
    /// 
    /// This application processes Excel files with delegate comments and uses ML.NET
    /// to classify comments into action categories (Add, Update, Term, Other).
    /// 
    /// Features:
    /// - Reads Excel files with multiple sheets
    /// - Processes "Delegate Comments" column using ML classification
    /// - Handles missing delegate comments (defaults to Add)
    /// - Creates separate output Excel files for each category
    /// - Includes comprehensive logging and error handling
    /// </summary>
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("=== TPDM Automation Application ===");
            Console.WriteLine("Processing Excel files with ML-based comment classification\n");

            try
            {
                // Check if user wants to generate sample data
                if (args.Length > 0 && args[0].Equals("--generate-sample", StringComparison.OrdinalIgnoreCase))
                {
                    GenerateSampleData();
                    return 0;
                }

                // Setup dependency injection and logging
                var serviceProvider = ConfigureServices();
                var logger = serviceProvider.GetRequiredService<ILogger<Program>>();
                var appService = serviceProvider.GetRequiredService<ApplicationService>();

                logger.LogInformation("Application started.");

                // Get configuration from command line arguments or use defaults
                var config = GetConfiguration(args, logger);
                
                if (!ValidateConfiguration(config, logger))
                {
                    return -1;
                }

                // Process the Excel file
                var success = await appService.ProcessExcelFileAsync(config);

                if (success)
                {
                    Console.WriteLine("\n✅ Processing completed successfully!");
                    Console.WriteLine($"📁 Output files created in: {config.OutputDirectory}");
                    logger.LogInformation("Application completed successfully.");
                    return 0;
                }
                else
                {
                    Console.WriteLine("\n❌ Processing failed. Check logs for details.");
                    logger.LogError("Application completed with errors.");
                    return -1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n💥 Unexpected error: {ex.Message}");
                Console.WriteLine("Check logs for detailed error information.");
                return -1;
            }
        }

        /// <summary>
        /// Generates sample Excel file for testing
        /// </summary>
        private static void GenerateSampleData()
        {
            try
            {
                var sampleFilePath = Path.Combine(Directory.GetCurrentDirectory(), "TestData", "sample_employee_data.xlsx");
                Directory.CreateDirectory(Path.GetDirectoryName(sampleFilePath)!);
                
                Console.WriteLine("Generating sample Excel file...");
                SampleDataGenerator.CreateSampleExcelFile(sampleFilePath);
                Console.WriteLine($"✅ Sample file created: {sampleFilePath}");
                Console.WriteLine("\nTo process this file, run:");
                Console.WriteLine($"dotnet run \"{sampleFilePath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error generating sample data: {ex.Message}");
            }
        }

        /// <summary>
        /// Configures dependency injection and services
        /// </summary>
        /// <returns>Configured service provider</returns>
        private static ServiceProvider ConfigureServices()
        {
            var services = new ServiceCollection();

            // Configure logging
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.SetMinimumLevel(LogLevel.Information);
            });

            // Register application services
            services.AddSingleton<MLService>();
            services.AddSingleton<ExcelService>();
            services.AddSingleton<ApplicationService>();

            return services.BuildServiceProvider();
        }

        /// <summary>
        /// Gets application configuration from command line arguments or prompts user
        /// </summary>
        /// <param name="args">Command line arguments</param>
        /// <param name="logger">Logger instance</param>
        /// <returns>Application configuration</returns>
        private static AppConfig GetConfiguration(string[] args, ILogger logger)
        {
            var config = new AppConfig();

            // Set default paths
            var currentDirectory = Directory.GetCurrentDirectory();
            config.TrainingDataPath = Path.Combine(currentDirectory, "Data", "training_data.csv");
            config.ModelPath = Path.Combine(currentDirectory, "Data", "comment_classifier.zip");

            if (args.Length >= 1)
            {
                // Input file path from command line
                config.InputFilePath = args[0];
            }
            else
            {
                // Prompt for input file path
                Console.Write("Enter the path to the Excel file to process: ");
                config.InputFilePath = Console.ReadLine()?.Trim() ?? "";
            }

            if (args.Length >= 2)
            {
                // Output directory from command line
                config.OutputDirectory = args[1];
            }
            else
            {
                // Use same directory as input file or prompt
                if (!string.IsNullOrEmpty(config.InputFilePath) && File.Exists(config.InputFilePath))
                {
                    config.OutputDirectory = Path.GetDirectoryName(config.InputFilePath) ?? currentDirectory;
                }
                else
                {
                    Console.Write("Enter the output directory (or press Enter for current directory): ");
                    var outputDir = Console.ReadLine()?.Trim();
                    config.OutputDirectory = string.IsNullOrEmpty(outputDir) ? currentDirectory : outputDir;
                }
            }

            logger.LogInformation($"Configuration:");
            logger.LogInformation($"  Input File: {config.InputFilePath}");
            logger.LogInformation($"  Output Directory: {config.OutputDirectory}");
            logger.LogInformation($"  Training Data: {config.TrainingDataPath}");
            logger.LogInformation($"  Model Path: {config.ModelPath}");

            return config;
        }

        /// <summary>
        /// Validates the application configuration
        /// </summary>
        /// <param name="config">Configuration to validate</param>
        /// <param name="logger">Logger instance</param>
        /// <returns>True if configuration is valid</returns>
        private static bool ValidateConfiguration(AppConfig config, ILogger logger)
        {
            var isValid = true;

            // Validate input file
            if (string.IsNullOrWhiteSpace(config.InputFilePath))
            {
                Console.WriteLine("❌ Input file path is required.");
                logger.LogError("Input file path is required.");
                isValid = false;
            }
            else if (!File.Exists(config.InputFilePath))
            {
                Console.WriteLine($"❌ Input file not found: {config.InputFilePath}");
                logger.LogError($"Input file not found: {config.InputFilePath}");
                isValid = false;
            }
            else if (!Path.GetExtension(config.InputFilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("❌ Input file must be an Excel file (.xlsx)");
                logger.LogError("Input file must be an Excel file (.xlsx)");
                isValid = false;
            }

            // Validate training data
            if (!File.Exists(config.TrainingDataPath))
            {
                Console.WriteLine($"❌ Training data file not found: {config.TrainingDataPath}");
                logger.LogError($"Training data file not found: {config.TrainingDataPath}");
                isValid = false;
            }

            // Validate/create output directory
            try
            {
                if (!Directory.Exists(config.OutputDirectory))
                {
                    Directory.CreateDirectory(config.OutputDirectory);
                    logger.LogInformation($"Created output directory: {config.OutputDirectory}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Cannot access output directory: {config.OutputDirectory}");
                logger.LogError(ex, $"Cannot access output directory: {config.OutputDirectory}");
                isValid = false;
            }

            return isValid;
        }
    }
}
