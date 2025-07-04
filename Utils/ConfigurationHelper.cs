using TPDMAutomation.Models;

namespace TPDMAutomation.Utils
{
    /// <summary>
    /// Utility class for handling application configuration
    /// </summary>
    public static class ConfigurationHelper
    {
        /// <summary>
        /// Creates a processing configuration from command line arguments
        /// </summary>
        /// <param name="args">Command line arguments</param>
        /// <returns>Processing configuration</returns>
        public static ProcessingConfig CreateFromArgs(string[] args)
        {
            var config = new ProcessingConfig();

            // Parse command line arguments
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "--input":
                    case "-i":
                        if (i + 1 < args.Length)
                        {
                            config.InputExcelPath = args[i + 1];
                            i++; // Skip next argument as it's the value
                        }
                        break;

                    case "--output":
                    case "-o":
                        if (i + 1 < args.Length)
                        {
                            config.OutputDirectory = args[i + 1];
                            i++; // Skip next argument as it's the value
                        }
                        break;

                    case "--column":
                    case "-c":
                        if (i + 1 < args.Length)
                        {
                            config.DelegateCommentsColumnName = args[i + 1];
                            i++; // Skip next argument as it's the value
                        }
                        break;

                    case "--default":
                    case "-d":
                        if (i + 1 < args.Length)
                        {
                            config.DefaultCategory = args[i + 1];
                            i++; // Skip next argument as it's the value
                        }
                        break;
                }
            }

            return config;
        }

        /// <summary>
        /// Prompts user for configuration if not provided via command line
        /// </summary>
        /// <param name="config">Current configuration</param>
        /// <returns>Updated configuration</returns>
        public static ProcessingConfig PromptForMissingConfig(ProcessingConfig config)
        {
            // Prompt for input Excel path if not provided
            if (string.IsNullOrWhiteSpace(config.InputExcelPath))
            {
                Console.Write("Enter the path to the input Excel file: ");
                config.InputExcelPath = Console.ReadLine() ?? "";
            }

            // Prompt for output directory if not provided
            if (string.IsNullOrWhiteSpace(config.OutputDirectory))
            {
                Console.Write("Enter the output directory path: ");
                config.OutputDirectory = Console.ReadLine() ?? "";
            }

            // Use defaults for other values if not provided
            if (string.IsNullOrWhiteSpace(config.DelegateCommentsColumnName))
            {
                config.DelegateCommentsColumnName = "Delegate Comments";
            }

            if (string.IsNullOrWhiteSpace(config.DefaultCategory))
            {
                config.DefaultCategory = "ADD";
            }

            return config;
        }

        /// <summary>
        /// Displays usage information for the application
        /// </summary>
        public static void ShowUsage()
        {
            Console.WriteLine("TPDM Automation - Excel Processing with ML Classification");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  TPDMAutomation --input <excel_file> --output <output_directory> [options]");
            Console.WriteLine();
            Console.WriteLine("Required Arguments:");
            Console.WriteLine("  --input, -i    Path to the input Excel file");
            Console.WriteLine("  --output, -o   Output directory for generated Excel files");
            Console.WriteLine();
            Console.WriteLine("Optional Arguments:");
            Console.WriteLine("  --column, -c   Name of the delegate comments column (default: 'Delegate Comments')");
            Console.WriteLine("  --default, -d  Default category when no comments found (default: 'ADD')");
            Console.WriteLine("  --help, -h     Show this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  TPDMAutomation -i \"C:\\data\\input.xlsx\" -o \"C:\\output\"");
            Console.WriteLine("  TPDMAutomation --input \"./data.xlsx\" --output \"./results\" --column \"Comments\" --default \"UPDATE\"");
            Console.WriteLine();
            Console.WriteLine("Description:");
            Console.WriteLine("  This application processes Excel files with multiple sheets, classifies delegate comments");
            Console.WriteLine("  using ML.NET into categories (Add, Update, Term, Other), and generates separate Excel");
            Console.WriteLine("  files for each category.");
        }

        /// <summary>
        /// Validates the configuration
        /// </summary>
        /// <param name="config">Configuration to validate</param>
        /// <returns>True if valid, false otherwise</returns>
        public static bool ValidateConfig(ProcessingConfig config)
        {
            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(config.InputExcelPath))
            {
                errors.Add("Input Excel file path is required");
            }

            if (string.IsNullOrWhiteSpace(config.OutputDirectory))
            {
                errors.Add("Output directory is required");
            }

            if (errors.Any())
            {
                Console.WriteLine("Configuration Errors:");
                foreach (var error in errors)
                {
                    Console.WriteLine($"  - {error}");
                }
                return false;
            }

            return true;
        }
    }
}