###############################################################################
# SVM Classification of Handmade Paper Products Using NIR HSI
#
# Two preprocessing pipelines:
#   (A) L2 normalization only
#   (B) 2nd derivative (Savitzky–Golay) + L2 normalization
#
# Classifier:
#   Support Vector Machine (RBF kernel, caret::svmRadial)
#
# Input:
#   NIR_Summary.xlsx (sheet = "CL_Product")
#
###############################################################################

library(openxlsx)
library(prospectr)
library(caret)
library(kernlab)

set.seed(123)

###############################################################################
# 1. Load Data -----------------------------------------------------------------
###############################################################################

data_file  <- "NIR_Summary.xlsx"     # e.g., "data/NIR_Summary.xlsx"
sheet_name <- "CL_Product"

raw <- read.xlsx(xlsxFile = data_file, sheet = sheet_name)

Sample <- as.factor(raw$Sample)
X      <- as.matrix(raw[, -1])     # spectral matrix

###############################################################################
# 2. L2 Normalization Function -------------------------------------------------
###############################################################################

euclidean_normalize <- function(x) {
  x / sqrt(sum(x^2))
}

###############################################################################
# 3. Preprocessing Pipelines ---------------------------------------------------
###############################################################################

# (A) L2 normalization only
X_L2 <- t(apply(X, 1, euclidean_normalize))
data_L2 <- data.frame(Sample = Sample, X_L2)

# (B) 2nd derivative + L2 normalization
X_2d <- t(apply(X, 1, function(x) savitzkyGolay(x, m = 2, p = 3, w = 7)))
X_2d_L2 <- t(apply(X_2d, 1, euclidean_normalize))
data_2dL2 <- data.frame(Sample = Sample, X_2d_L2)

###############################################################################
# 4. SVM Training & Evaluation Function ----------------------------------------
###############################################################################

evaluate_svm <- function(df) {
  
  stopifnot(is.factor(df$Sample))
  
  # ---- Stratified 70/30 split ----
  idx   <- createDataPartition(df$Sample, p = 0.70, list = FALSE)
  train <- df[idx, ]
  test  <- df[-idx, ]
  
  # ---- Hyperparameter grid ----
  C_grid     <- 10^(0:5)       # 1, 10, 1e2, ..., 1e5
  gamma_grid <- 10^-(1:6)      # 1e−1 ... 1e−6  (caret uses sigma = gamma)
  
  tune_grid <- expand.grid(C = C_grid, sigma = gamma_grid)
  
  # ---- Cross-validation ----
  ctrl <- trainControl(
    method = "repeatedcv",
    number = 3,
    repeats = 1,
    summaryFunction = defaultSummary,
    savePredictions = "final",
    allowParallel = TRUE
  )
  
  # ---- Train SVM-RBF ----
  svm_fit <- train(
    Sample ~ .,
    data      = train,
    method    = "svmRadial",
    tuneGrid  = tune_grid,
    metric    = "Accuracy",
    trControl = ctrl
  )
  
  # Print best parameters
  best <- svm_fit$bestTune
  
  # ---- Evaluate on test set ----
  pred <- predict(svm_fit, newdata = test)
  cm   <- confusionMatrix(pred, test$Sample)
  
  # ---- Macro-F1 ----
  byc <- cm$byClass
  f1  <- if (is.matrix(byc)) byc[, "F1"] else byc["F1"]
  macro_F1 <- mean(as.numeric(f1), na.rm = TRUE)
  
  list(
    best_C      = best$C,
    best_sigma  = best$sigma,
    accuracy    = unname(cm$overall["Accuracy"]),
    macro_F1    = macro_F1,
    confusion   = cm$table
  )
}

###############################################################################
# 5. Run Models for Both Preprocessing Methods ---------------------------------
###############################################################################

svm_L2    <- evaluate_svm(data_L2)
svm_2dL2  <- evaluate_svm(data_2dL2)

###############################################################################
# 6. Output Summary -------------------------------------------------------------
###############################################################################

cat("\n==================== SVM Classification Results ====================\n")

cat("\n--- (A) L2 NORMALIZATION ONLY ---\n")
print(svm_L2[c("best_C", "best_sigma", "accuracy", "macro_F1")])

cat("\n--- (B) 2ND DERIVATIVE + L2 NORMALIZATION ---\n")
print(svm_2dL2[c("best_C", "best_sigma", "accuracy", "macro_F1")])

cat("\n====================================================================\n")
