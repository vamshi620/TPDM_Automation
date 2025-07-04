using Microsoft.ML.Data;

namespace TPDMAutomation.Models
{
    /// <summary>
    /// Data model for training the ML.NET text classification model
    /// </summary>
    public class CommentData
    {
        /// <summary>
        /// The delegate comment text to classify
        /// </summary>
        [LoadColumn(0)]
        public string Comment { get; set; } = string.Empty;

        /// <summary>
        /// The predicted label (Add, Update, Term, Other)
        /// </summary>
        [LoadColumn(1)]
        public string Label { get; set; } = string.Empty;
    }

    /// <summary>
    /// Prediction result from the ML model
    /// </summary>
    public class CommentPrediction
    {
        /// <summary>
        /// Predicted label for the comment
        /// </summary>
        [ColumnName("PredictedLabel")]
        public string PredictedLabel { get; set; } = string.Empty;

        /// <summary>
        /// Confidence score for the prediction
        /// </summary>
        [ColumnName("Score")]
        public float[] Score { get; set; } = Array.Empty<float>();
    }
}