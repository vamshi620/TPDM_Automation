namespace TPDMAutomation.Models
{
    /// <summary>
    /// Represents data from an Excel row with prediction
    /// </summary>
    public class ExcelRowData
    {
        /// <summary>
        /// The original row index in the Excel sheet
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// The delegate comment text
        /// </summary>
        public string DelegateComment { get; set; } = string.Empty;

        /// <summary>
        /// The predicted category (Add, Update, Term, Other)
        /// </summary>
        public string PredictedCategory { get; set; } = string.Empty;

        /// <summary>
        /// All column values from the original row
        /// </summary>
        public Dictionary<string, object> ColumnValues { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// The sheet name this row belongs to
        /// </summary>
        public string SheetName { get; set; } = string.Empty;
    }

    /// <summary>
    /// Configuration for processing Excel files
    /// </summary>
    public class ProcessingConfig
    {
        /// <summary>
        /// Path to the input Excel file
        /// </summary>
        public string InputExcelPath { get; set; } = string.Empty;

        /// <summary>
        /// Output directory for generated Excel files
        /// </summary>
        public string OutputDirectory { get; set; } = string.Empty;

        /// <summary>
        /// Column name to look for delegate comments
        /// </summary>
        public string DelegateCommentsColumnName { get; set; } = "Delegate Comments";

        /// <summary>
        /// Default category when no delegate comments column is found
        /// </summary>
        public string DefaultCategory { get; set; } = "Add";
    }
}