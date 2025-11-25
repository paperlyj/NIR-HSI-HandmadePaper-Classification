###############################################################################
# Feedforward Neural Network (FNN) Classification of Handmade Paper Products
#
# Two preprocessing pipelines:
#   (A) L2 normalization only
#   (B) 2nd derivative (Savitzky–Golay) + L2 normalization
#
# Classifier:
#   Feedforward neural network (ANN2::neuralnetwork)
#
# Input:
#   NIR_Summary.xlsx (sheet = "CL_Product")
###############################################################################

library(openxlsx)
library(prospectr)
library(caret)
library(ANN2)

set.seed(1234)

###############################################################################
# 1. Load Data -----------------------------------------------------------------
###############################################################################

data_file  <- "NIR_Summary.xlsx"    # e.g., "data/NIR_Summary.xlsx"
sheet_name <- "CL_Product"

raw <- read.xlsx(xlsxFile = data_file, sheet = sheet_name)

Sample <- as.factor(raw$Sample)
X      <- as.matrix(raw[, -1])   # spectral matrix only

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
# 4. FNN Training & Evaluation Function ----------------------------------------
###############################################################################

evaluate_fnn <- function(df,
                         hidden_layers = c(16),
                         lr            = 0.01,
                         epochs        = 300,
                         val_prop      = 0.3) {
  
  stopifnot(is.factor(df$Sample))
  
  # ---- Stratified 70/30 split ----
  idx   <- createDataPartition(df$Sample, p = 0.70, list = FALSE)
  train <- df[idx, ]
  test  <- df[-idx, ]
  
  x_train <- as.matrix(train[, -1])
  y_train <- train[, 1]
  
  x_test  <- as.matrix(test[, -1])
  y_test  <- test[, 1]
  
  # ---- Train FNN ----
  fnn_fit <- neuralnetwork(
    X              = x_train,
    y              = y_train,
    hidden.layers  = hidden_layers,
    regression     = FALSE,
    standardize    = TRUE,
    loss.type      = "log",
    activ.functions = "relu",
    optim.type     = "adam",
    learn.rates    = lr,
    val.prop       = val_prop,
    n.epochs       = epochs
  )
  
  # ---- Prediction ----
  pred <- predict(fnn_fit, newdata = x_test)
  pred_class <- as.factor(pred$predictions)
  true_class <- as.factor(y_test)
  
  cm <- confusionMatrix(pred_class, reference = true_class,
                        mode = "everything")
  
  # ---- Macro-F1 ----
  byc    <- cm$byClass
  f1_vec <- if (is.matrix(byc)) byc[, "F1"] else byc["F1"]
  macro_F1 <- mean(as.numeric(f1_vec), na.rm = TRUE)
  
  list(
    accuracy   = unname(cm$overall["Accuracy"]),
    macro_F1   = macro_F1,
    confusion  = cm$table
  )
}

###############################################################################
# 5. Run FNN Under Both Preprocessing Settings ---------------------------------
###############################################################################

fnn_L2    <- evaluate_fnn(data_L2,
                          hidden_layers = c(16),
                          lr            = 0.01,
                          epochs        = 300)

fnn_2dL2  <- evaluate_fnn(data_2dL2,
                          hidden_layers = c(16),
                          lr            = 0.01,
                          epochs        = 300)

###############################################################################
# 6. Output Summary -------------------------------------------------------------
###############################################################################

cat("\n==================== FNN Classification Results ====================\n")

cat("\n--- (A) L2 NORMALIZATION ONLY ---\n")
cat(sprintf("Accuracy : %.4f\n", fnn_L2$accuracy))
cat(sprintf("Macro F1 : %.4f\n", fnn_L2$macro_F1))

cat("\n--- (B) 2ND DERIVATIVE + L2 NORMALIZATION ---\n")
cat(sprintf("Accuracy : %.4f\n", fnn_2dL2$accuracy))
cat(sprintf("Macro F1 : %.4f\n", fnn_2dL2$macro_F1))

cat("\n====================================================================\n")
