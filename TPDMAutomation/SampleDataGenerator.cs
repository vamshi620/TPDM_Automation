using OfficeOpenXml;

namespace TPDMAutomation
{
    /// <summary>
    /// Helper class to create sample Excel files for testing
    /// </summary>
    public static class SampleDataGenerator
    {
        /// <summary>
        /// Creates a sample Excel file with test data
        /// </summary>
        public static void CreateSampleExcelFile(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();

            // Create first sheet
            var sheet1 = package.Workbook.Worksheets.Add("Employees");
            
            // Headers
            sheet1.Cells[1, 1].Value = "Employee ID";
            sheet1.Cells[1, 2].Value = "Employee Name";
            sheet1.Cells[1, 3].Value = "Department";
            sheet1.Cells[1, 4].Value = "Delegate Comments";
            sheet1.Cells[1, 5].Value = "Manager";

            // Sample data with various comment types
            var sampleData = new object[,]
            {
                { "EMP001", "John Smith", "IT", "New employee starting next month", "Jane Doe" },
                { "EMP002", "Mary Johnson", "HR", "Employee information needs updating", "Bob Wilson" },
                { "EMP003", "David Brown", "Finance", "Employee is leaving the company", "Alice Cooper" },
                { "EMP004", "Sarah Davis", "Marketing", "General inquiry about employee", "Tom Jones" },
                { "EMP005", "Mike Wilson", "IT", "Hiring new team member for the project", "Jane Doe" },
                { "EMP006", "Lisa Anderson", "HR", "Termination effective immediately", "Bob Wilson" },
                { "EMP007", "Robert Taylor", "Sales", "", "Carol White" }, // Empty comment - should default to Add
                { "EMP008", "Emily Clark", "Finance", "Change in employee status", "Alice Cooper" },
                { "EMP009", "James Lewis", "Marketing", "End of contract", "Tom Jones" },
                { "EMP010", "Jennifer Miller", "IT", "Review pending for employee", "Jane Doe" }
            };

            // Fill data
            for (int i = 0; i < sampleData.GetLength(0); i++)
            {
                for (int j = 0; j < sampleData.GetLength(1); j++)
                {
                    sheet1.Cells[i + 2, j + 1].Value = sampleData[i, j];
                }
            }

            // Create second sheet
            var sheet2 = package.Workbook.Worksheets.Add("Contractors");
            
            // Headers
            sheet2.Cells[1, 1].Value = "Contractor ID";
            sheet2.Cells[1, 2].Value = "Contractor Name";
            sheet2.Cells[1, 3].Value = "Project";
            sheet2.Cells[1, 4].Value = "Delegate Comments";
            sheet2.Cells[1, 5].Value = "Start Date";

            // Sample contractor data
            var contractorData = new object[,]
            {
                { "CON001", "Alex Rodriguez", "Project Alpha", "Adding additional resource to the team", "2024-01-15" },
                { "CON002", "Maria Garcia", "Project Beta", "Update contact information", "2024-02-01" },
                { "CON003", "Thomas Kim", "Project Gamma", "Contract conclusion", "2024-03-10" },
                { "CON004", "Sophie Turner", "Project Delta", "", "2024-01-20" }, // Empty comment
                { "CON005", "Chris Evans", "Project Echo", "Fresh recruit for the position", "2024-02-15" }
            };

            // Fill contractor data
            for (int i = 0; i < contractorData.GetLength(0); i++)
            {
                for (int j = 0; j < contractorData.GetLength(1); j++)
                {
                    sheet2.Cells[i + 2, j + 1].Value = contractorData[i, j];
                }
            }

            // Create third sheet without "Delegate Comments" column
            var sheet3 = package.Workbook.Worksheets.Add("Interns");
            
            // Headers (no delegate comments column)
            sheet3.Cells[1, 1].Value = "Intern ID";
            sheet3.Cells[1, 2].Value = "Intern Name";
            sheet3.Cells[1, 3].Value = "University";
            sheet3.Cells[1, 4].Value = "Mentor";

            // Sample intern data
            var internData = new object[,]
            {
                { "INT001", "Sam Parker", "MIT", "John Smith" },
                { "INT002", "Rachel Green", "Stanford", "Mary Johnson" },
                { "INT003", "Kevin Scott", "Harvard", "David Brown" }
            };

            // Fill intern data
            for (int i = 0; i < internData.GetLength(0); i++)
            {
                for (int j = 0; j < internData.GetLength(1); j++)
                {
                    sheet3.Cells[i + 2, j + 1].Value = internData[i, j];
                }
            }

            // Auto-fit columns for all sheets
            sheet1.Cells.AutoFitColumns();
            sheet2.Cells.AutoFitColumns();
            sheet3.Cells.AutoFitColumns();

            // Format headers
            foreach (var sheet in package.Workbook.Worksheets)
            {
                for (int col = 1; col <= sheet.Dimension.Columns; col++)
                {
                    sheet.Cells[1, col].Style.Font.Bold = true;
                    sheet.Cells[1, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }
            }

            package.SaveAs(new FileInfo(filePath));
        }
    }
}