###############################################################################
# KNN Classification of Traditional Handmade Papers (Country Level)
# - (A) L2 normalization only
# - (B) Savitzky–Golay 2nd derivative + L2 normalization
# - Train/test split + cross-validated KNN
# - Mean spectra plots for each country (raw & 2nd derivative)
#
# Input : NIR_Summary.xlsx (sheet = "CL_Product")
###############################################################################

## 0. Packages and global options ---------------------------------------------

library(openxlsx)
library(prospectr)
library(caret)
library(kknn)

set.seed(123)

## 1. Load data ----------------------------------------------------------------

# Path to the NIR summary file (place in project root or /data/)
data_file  <- "NIR_Summary.xlsx"    # e.g. "data/NIR_Summary.xlsx"
sheet_name <- "CL_Country"

dat <- read.xlsx(xlsxFile = data_file, sheet = sheet_name)

Sample <- as.factor(dat$Sample)     # class label (e.g. China / Japan / Korea)
X      <- as.matrix(dat[, -1])      # spectral matrix only

# Extract numeric wavelength values from column names
wl_char <- colnames(dat[, -1])
Wavelength <- as.numeric(gsub("^X", "", wl_char))

## 2. L2 normalization function ------------------------------------------------

euclidean_normalize <- function(x) {
  x / sqrt(sum(x^2))
}

## 3. Preprocessing pipelines --------------------------------------------------

# (A) L2 normalization only
X_L2 <- t(apply(X, 1, euclidean_normalize))
data_L2 <- data.frame(Sample = Sample, X_L2)

# (B) Savitzky–Golay 2nd derivative + L2 normalization
X_2d <- t(apply(X, 1, function(x) savitzkyGolay(x, m = 2, p = 3, w = 7)))
X_2d_L2 <- t(apply(X_2d, 1, euclidean_normalize))
data_2dL2 <- data.frame(Sample = Sample, X_2d_L2)

## 4. Helper function: train and evaluate KNN ----------------------------------

evaluate_knn <- function(df) {
  stopifnot(is.factor(df$Sample))
  
  # Stratified train/test split (70/30)
  idx   <- createDataPartition(df$Sample, p = 0.7, list = FALSE)
  train <- df[idx, ]
  test  <- df[-idx, ]
  
  # Candidate k values
  k_grid   <- seq(1, 9, 2)
  tune_grid <- expand.grid(k = k_grid)
  
  ctrl <- trainControl(
    method          = "repeatedcv",
    number          = 3,
    repeats         = 1,
    summaryFunction = defaultSummary,
    savePredictions = "final",
    allowParallel   = TRUE
  )
  
  knn_fit <- caret::train(
    Sample ~ .,
    data      = train,
    method    = "knn",
    preProcess = c("center", "scale"),
    tuneGrid   = tune_grid,
    metric     = "Accuracy",
    trControl  = ctrl
  )
  
  pred <- predict(knn_fit, newdata = test)
  cm   <- caret::confusionMatrix(pred, test$Sample)
  
  byc      <- cm$byClass
  f1_vec   <- if (is.matrix(byc)) byc[, "F1"] else byc["F1"]
  macro_F1 <- mean(as.numeric(f1_vec), na.rm = TRUE)
  
  list(
    best_k   = knn_fit$bestTune$k,
    accuracy = unname(cm$overall["Accuracy"]),
    macro_F1 = macro_F1
  )
}

## 5. Run KNN under both preprocessing settings --------------------------------

result_L2   <- evaluate_knn(data_L2)
result_2dL2 <- evaluate_knn(data_2dL2)

cat("\n===== KNN Model Comparison (Country Classification) =====\n")
cat("L2 only:\n")
print(result_L2)

cat("\n2nd derivative + L2:\n")
print(result_2dL2)