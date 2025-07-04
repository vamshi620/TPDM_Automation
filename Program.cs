using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using TPDMAutomation.Services;
using TPDMAutomation.Utils;
using ClosedXML.Excel;

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
                Console.WriteLine("Generating test data for TPDM Automation...");
                
                // Create Test directory
                var testDir = Path.Combine(Directory.GetCurrentDirectory(), "Test");
                Directory.CreateDirectory(testDir);
                
                var inputFile = Path.Combine(testDir, "sample_input.xlsx");
                
                Console.WriteLine($"Creating sample Excel file: {inputFile}");
                
                using (var workbook = new XLWorkbook())
                {
                    // Sheet 1: HR Records with Delegate Comments
                    var sheet1 = workbook.Worksheets.Add("HR_Records");
                    CreateSheetWithComments(sheet1, "HR Records");
                    
                    // Sheet 2: Employee Data with Delegate Comments  
                    var sheet2 = workbook.Worksheets.Add("Employee_Data");
                    CreateSheetWithComments(sheet2, "Employee Data");
                    
                    // Sheet 3: Payroll Info without Delegate Comments (should default to ADD)
                    var sheet3 = workbook.Worksheets.Add("Payroll_Info");
                    CreateSheetWithoutComments(sheet3, "Payroll Info");
                    
                    workbook.SaveAs(inputFile);
                }
                
                Console.WriteLine("Test data generation completed successfully!");
                Console.WriteLine($"Sample input file created: {inputFile}");
                Console.WriteLine("You can now test the application using:");
                Console.WriteLine($"dotnet run -- -i \"{inputFile}\" -o \"{testDir}\"");
                
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating test data: {ex.Message}");
                return 1;
            }
        }
        
        /// <summary>
        /// Creates a worksheet with delegate comments column
        /// </summary>
        /// <param name="worksheet">The worksheet to populate</param>
        /// <param name="sheetType">Type of sheet for varied data</param>
        private static void CreateSheetWithComments(IXLWorksheet worksheet, string sheetType)
        {
            // Headers
            worksheet.Cell(1, 1).Value = "Employee ID";
            worksheet.Cell(1, 2).Value = "Employee Name";
            worksheet.Cell(1, 3).Value = "Department";
            worksheet.Cell(1, 4).Value = "Position";
            worksheet.Cell(1, 5).Value = "Delegate Comments";
            worksheet.Cell(1, 6).Value = "Date";
            
            // Sample comments that should map to different categories
            var testComments = new[]
            {
                ("E001", "John Smith", "IT", "Developer", "Added new employee to the system", "Add"),
                ("E002", "Jane Doe", "HR", "Manager", "Updated employee information", "Update"), 
                ("E003", "Bob Johnson", "Finance", "Analyst", "Employee has been terminated", "Term"),
                ("E004", "Alice Brown", "Marketing", "Coordinator", "Review pending for this employee", "Other"),
                ("E005", "Charlie Wilson", "IT", "Senior Developer", "New hire has been processed", "Add"),
                ("E006", "Diana Clark", "HR", "Specialist", "Modified contact details", "Update"),
                ("E007", "Edward Lee", "Finance", "Director", "Terminated due to policy violation", "Term"),
                ("E008", "Fiona Davis", "Marketing", "Manager", "Under investigation", "Other"),
                ("E009", "George Miller", "IT", "Architect", "Onboarding new staff member", "Add"),
                ("E010", "Helen Taylor", "HR", "Assistant", "Changed department assignment", "Update")
            };
            
            // Add data rows
            for (int i = 0; i < testComments.Length; i++)
            {
                var row = i + 2;
                var comment = testComments[i];
                
                worksheet.Cell(row, 1).Value = comment.Item1;
                worksheet.Cell(row, 2).Value = comment.Item2;
                worksheet.Cell(row, 3).Value = comment.Item3;
                worksheet.Cell(row, 4).Value = comment.Item4;
                worksheet.Cell(row, 5).Value = comment.Item5;
                worksheet.Cell(row, 6).Value = DateTime.Now.AddDays(-i).ToString("yyyy-MM-dd");
            }
            
            // Format the headers
            var headerRange = worksheet.Range(1, 1, 1, 6);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            
            // Auto-fit columns
            worksheet.Columns().AdjustToContents();
        }
        
        /// <summary>
        /// Creates a worksheet without delegate comments column (should default to ADD)
        /// </summary>
        /// <param name="worksheet">The worksheet to populate</param>
        /// <param name="sheetType">Type of sheet for context</param>
        private static void CreateSheetWithoutComments(IXLWorksheet worksheet, string sheetType)
        {
            // Headers - no "Delegate Comments" column
            worksheet.Cell(1, 1).Value = "Employee ID";
            worksheet.Cell(1, 2).Value = "Employee Name";
            worksheet.Cell(1, 3).Value = "Salary";
            worksheet.Cell(1, 4).Value = "Benefits";
            worksheet.Cell(1, 5).Value = "Tax Code";
            
            // Sample data
            var payrollData = new[]
            {
                ("P001", "Michael Green", "$75,000", "Health+Dental", "TC001"),
                ("P002", "Sarah White", "$82,000", "Health+Dental+Vision", "TC002"),
                ("P003", "David Black", "$68,000", "Health", "TC001"),
                ("P004", "Lisa Gray", "$95,000", "Full Package", "TC003"),
                ("P005", "Tom Blue", "$71,000", "Health+Dental", "TC001")
            };
            
            // Add data rows
            for (int i = 0; i < payrollData.Length; i++)
            {
                var row = i + 2;
                var data = payrollData[i];
                
                worksheet.Cell(row, 1).Value = data.Item1;
                worksheet.Cell(row, 2).Value = data.Item2;
                worksheet.Cell(row, 3).Value = data.Item3;
                worksheet.Cell(row, 4).Value = data.Item4;
                worksheet.Cell(row, 5).Value = data.Item5;
            }
            
            // Format the headers
            var headerRange = worksheet.Range(1, 1, 1, 5);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            
            // Auto-fit columns
            worksheet.Columns().AdjustToContents();
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
