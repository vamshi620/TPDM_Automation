using Microsoft.ML.Data;

namespace TPDMAutomation.Models
{
    /// <summary>
    /// Represents training data for the ML model
    /// Contains comment text and the corresponding action category
    /// </summary>
    public class CommentData
    {
        [LoadColumn(0)]
        public string Comment { get; set; } = string.Empty;

        [LoadColumn(1)]
        public string Action { get; set; } = string.Empty;
    }

    /// <summary>
    /// Represents the prediction result from the ML model
    /// </summary>
    public class CommentPrediction
    {
        [ColumnName("PredictedLabel")]
        public string PredictedAction { get; set; } = string.Empty;

        [ColumnName("Score")]
        public float[] Score { get; set; } = Array.Empty<float>();

        [ColumnName("Label")]
        public uint Label { get; set; }
    }

    /// <summary>
    /// Represents a row of data from Excel with prediction result
    /// </summary>
    public class ExcelRowData
    {
        public int RowNumber { get; set; }
        public string DelegateComment { get; set; } = string.Empty;
        public string PredictedAction { get; set; } = string.Empty;
        public Dictionary<string, object> OtherColumns { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// Represents configuration for the application
    /// </summary>
    public class AppConfig
    {
        public string InputFilePath { get; set; } = string.Empty;
        public string OutputDirectory { get; set; } = string.Empty;
        public string TrainingDataPath { get; set; } = string.Empty;
        public string ModelPath { get; set; } = string.Empty;
    }
}