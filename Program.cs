using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using TPDMAutomation.Services;
using TPDMAutomation.Utils;

namespace TPDMAutomation
{
    /// <summary>
    /// Main program class for TPDM Automation application
    /// </summary>
    class Program
    {
        /// <summary>
        /// Main entry point of the application
        /// </summary>
        /// <param name="args">Command line arguments</param>
        /// <returns>Exit code (0 for success, 1 for error)</returns>
        static async Task<int> Main(string[] args)
        {
            try
            {
                // Check for help argument
                if (args.Length == 0 || args.Contains("--help") || args.Contains("-h"))
                {
                    ConfigurationHelper.ShowUsage();
                    return 0;
                }

                // Check for generate test data argument
                if (args.Contains("--generate-test-data"))
                {
                    return GenerateTestData();
                }

                // Setup dependency injection and logging
                var serviceProvider = ConfigureServices();
                var logger = serviceProvider.GetRequiredService<ILogger<Program>>();

                logger.LogInformation("TPDM Automation Application Starting...");

                // Parse configuration from command line arguments
                var config = ConfigurationHelper.CreateFromArgs(args);
                config = ConfigurationHelper.PromptForMissingConfig(config);

                // Validate configuration
                if (!ConfigurationHelper.ValidateConfig(config))
                {
                    logger.LogError("Invalid configuration provided");
                    return 1;
                }

                // Display configuration
                logger.LogInformation("Configuration:");
                logger.LogInformation("  Input Excel File: {InputPath}", config.InputExcelPath);
                logger.LogInformation("  Output Directory: {OutputDirectory}", config.OutputDirectory);
                logger.LogInformation("  Delegate Comments Column: {ColumnName}", config.DelegateCommentsColumnName);
                logger.LogInformation("  Default Category: {DefaultCategory}", config.DefaultCategory);

                // Get the main application service and process the Excel file
                var appService = serviceProvider.GetRequiredService<TPDMApplicationService>();
                bool success = await appService.ProcessExcelFileAsync(config);

                if (success)
                {
                    logger.LogInformation("TPDM Automation completed successfully!");
                    Console.WriteLine("\nProcessing completed successfully!");
                    Console.WriteLine($"Output files have been created in: {config.OutputDirectory}");
                    return 0;
                }
                else
                {
                    logger.LogError("TPDM Automation failed");
                    Console.WriteLine("\nProcessing failed. Please check the logs for details.");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
                Console.WriteLine("Use --help for usage information.");
                return 1;
            }
        }

        /// <summary>
        /// Generates test data for the application
        /// </summary>
        /// <returns>Exit code</returns>
        private static int GenerateTestData()
        {
            try
            {
                Console.WriteLine("Test data generation is not implemented yet.");
                Console.WriteLine("Please provide your own Excel file with the required format.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating test data: {ex.Message}");
                return 1;
            }
        }

        /// <summary>
        /// Configures dependency injection services
        /// </summary>
        /// <returns>Service provider</returns>
        private static ServiceProvider ConfigureServices()
        {
            var services = new ServiceCollection();

            // Add logging
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.SetMinimumLevel(LogLevel.Information);
            });

            // Add application services
            services.AddTransient<MLModelService>();
            services.AddTransient<ExcelProcessingService>();
            services.AddTransient<TPDMApplicationService>();

            return services.BuildServiceProvider();
        }
    }
}
