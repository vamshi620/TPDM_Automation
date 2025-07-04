using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.Extensions.Logging;
using TPDMAutomation.Models;

namespace TPDMAutomation.Services
{
    /// <summary>
    /// Service for training and using ML.NET model for comment classification
    /// </summary>
    public class MLModelService
    {
        private readonly MLContext _mlContext;
        private readonly ILogger<MLModelService> _logger;
        private ITransformer? _trainedModel;

        /// <summary>
        /// Initializes a new instance of the MLModelService
        /// </summary>
        /// <param name="logger">Logger instance</param>
        public MLModelService(ILogger<MLModelService> logger)
        {
            _mlContext = new MLContext(seed: 0);
            _logger = logger;
        }

        /// <summary>
        /// Trains the ML model using the provided training data
        /// </summary>
        /// <param name="trainingDataPath">Path to the training CSV file</param>
        /// <returns>True if training was successful, false otherwise</returns>
        public bool TrainModel(string trainingDataPath)
        {
            try
            {
                _logger.LogInformation("Starting ML model training with data from: {TrainingDataPath}", trainingDataPath);

                // Load training data
                IDataView dataView = _mlContext.Data.LoadFromTextFile<CommentData>(
                    path: trainingDataPath,
                    hasHeader: true,
                    separatorChar: ',');

                _logger.LogInformation("Training data loaded successfully");

                // Define the training pipeline
                var pipeline = _mlContext.Transforms.Conversion.MapValueToKey(inputColumnName: "Label", outputColumnName: "Label")
                    .Append(_mlContext.Transforms.Text.FeaturizeText(outputColumnName: "Features", inputColumnName: "Comment"))
                    .Append(_mlContext.MulticlassClassification.Trainers.SdcaMaximumEntropy(labelColumnName: "Label", featureColumnName: "Features"))
                    .Append(_mlContext.Transforms.Conversion.MapKeyToValue("PredictedLabel"));

                _logger.LogInformation("Training pipeline created");

                // Train the model
                _trainedModel = pipeline.Fit(dataView);

                _logger.LogInformation("Model training completed successfully");

                // Evaluate the model
                var predictions = _trainedModel.Transform(dataView);
                var metrics = _mlContext.MulticlassClassification.Evaluate(predictions, labelColumnName: "Label");

                _logger.LogInformation("Model Evaluation Metrics:");
                _logger.LogInformation("MicroAccuracy: {MicroAccuracy:0.###}", metrics.MicroAccuracy);
                _logger.LogInformation("MacroAccuracy: {MacroAccuracy:0.###}", metrics.MacroAccuracy);

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred during model training");
                return false;
            }
        }

        /// <summary>
        /// Predicts the category for a given comment
        /// </summary>
        /// <param name="comment">The comment text to classify</param>
        /// <returns>The predicted category (Add, Update, Term, Other)</returns>
        public string PredictCategory(string comment)
        {
            if (_trainedModel == null)
            {
                _logger.LogWarning("Model is not trained. Cannot make prediction for comment: {Comment}", comment);
                return "Other";
            }

            try
            {
                // Create prediction engine
                var predictionEngine = _mlContext.Model.CreatePredictionEngine<CommentData, CommentPrediction>(_trainedModel);

                // Make prediction
                var prediction = predictionEngine.Predict(new CommentData { Comment = comment });

                _logger.LogDebug("Predicted category for comment '{Comment}': {PredictedLabel}", comment, prediction.PredictedLabel);

                return prediction.PredictedLabel ?? "Other";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred during prediction for comment: {Comment}", comment);
                return "Other";
            }
        }

        /// <summary>
        /// Checks if the model is trained and ready for predictions
        /// </summary>
        /// <returns>True if model is trained, false otherwise</returns>
        public bool IsModelTrained()
        {
            return _trainedModel != null;
        }
    }
}