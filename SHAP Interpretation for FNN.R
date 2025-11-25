###############################################################################
# SHAP Interpretation for Feedforward Neural Network (ANN2)
# NIR Hyperspectral Data – Country Classification
#
# This script:
#   1. Loads NIR HSI dataset (S_Country)
#   2. Applies L2 normalization
#   3. Trains a Feedforward Neural Network (ANN2)
#   4. Computes SHAP values using fastshap
#   5. Visualizes SHAP importance:
#        (A) Individual sample SHAP (image plot)
#        (B) Mean absolute SHAP + average spectra overlay
#
# Input:
#   NIR_Summary.xlsx (sheet = "S_Country")
###############################################################################

library(openxlsx)
library(prospectr)
library(caret)
library(ANN2)
library(fastshap)

set.seed(1234)

###############################################################################
# 1. Load data -----------------------------------------------------------------
###############################################################################

data_file  <- "NIR_Summary.xlsx"     # e.g., "data/NIR_Summary.xlsx"
sheet_name <- "S_Country"

df <- read.xlsx(data_file, sheet = sheet_name)
df$Sample <- as.factor(df$Sample)

Wavelength <- as.numeric(colnames(df[, -1]))
X <- as.matrix(df[, -1])

###############################################################################
# 2. Row-wise L2 normalization -------------------------------------------------
###############################################################################

row_norm <- sqrt(rowSums(X^2))
row_norm[row_norm == 0] <- 1e-12
X_l2 <- sweep(X, 1, row_norm, "/")

df[, -1] <- as.data.frame(X_l2)

###############################################################################
# 3. Train/Test Split -----------------------------------------------------------
###############################################################################

idx <- createDataPartition(df$Sample, p = 0.7, list = FALSE)
train <- df[idx, ]
test  <- df[-idx, ]

x_train <- as.matrix(train[, -1])
y_train <- train[, 1]

x_test  <- as.matrix(test[, -1])
y_test  <- test[, 1]

###############################################################################
# 4. Train Feedforward Neural Network (ANN2) -----------------------------------
###############################################################################

fnn_model <- neuralnetwork(
  X              = x_train,
  y              = y_train,
  hidden.layers  = c(16),
  regression     = FALSE,
  standardize    = TRUE,
  loss.type      = "log",
  activ.functions = "relu",
  optim.type     = "adam",
  learn.rates    = 0.01,
  val.prop       = 0.3,
  n.epochs       = 300
)

###############################################################################
# 5. SHAP Computation -----------------------------------------------------------
###############################################################################

pred_wrapper <- function(object, newdata) {
  pred <- predict(object, as.matrix(newdata))
  as.numeric(pred$probabilities[, 1])   # class 1 probability
}

shap_values <- fastshap::explain(
  object       = fnn_model,
  X            = as.data.frame(x_test),
  pred_wrapper = pred_wrapper,
  nsim         = 50,
  adjust       = TRUE
)

rownames(shap_values) <- as.character(y_test)

###############################################################################
# 6A. SHAP Visualization – Single Sample (Image Plot) --------------------------
###############################################################################

plot_single_shap <- function(sample_name) {
  
  shap_one <- data.frame(
    wavelength = as.numeric(colnames(x_test)),
    shap_value = as.numeric(shap_values[sample_name, ])
  )
  
  colmap <- colorRampPalette(c("royalblue", "white", "brown1"))(200)
  
  windows(7, 1.5)
  par(mar = c(1, 2, 1, 2))
  
  image(
    x = shap_one$wavelength,
    y = 1,
    z = t(matrix(shap_one$shap_value, nrow = 1)),
    col = colmap,
    xlab = "Wavelength (nm)", ylab = "",
    axes = FALSE,
    cex.lab = 1.5,
    xlim = c(min(Wavelength), max(Wavelength))
  )
  box(lwd = 3)
}

# Example:
# plot_single_shap("Korea")

###############################################################################
# 6B. SHAP Visualization – Mean SHAP + Average Spectrum Overlay ----------------
###############################################################################

# --- Compute class-wise average spectra ----
split_avg <- split(df[, -1], df$Sample)
avg_spec  <- sapply(split_avg, colMeans)
avg_spec  <- data.frame(Wavelength = Wavelength, avg_spec)

# --- Mean absolute SHAP over all test samples ----
mean_abs_shap <- apply(abs(shap_values), 2, mean)

shap_df <- data.frame(
  wavelength     = as.numeric(colnames(x_test)),
  mean_abs_shap  = mean_abs_shap
)

plot_mean_shap <- function() {
  
  colmap <- colorRampPalette(c("royalblue", "white", "brown1"))(200)
  
  windows(7, 4)
  par(mar = c(4, 2, 2, 2))
  
  image(
    x = shap_df$wavelength,
    y = 1,
    z = t(matrix(shap_df$mean_abs_shap, nrow = 1)),
    col = colmap,
    xlab = expression(bold("Wavelength (nm)")),
    ylab = "",
    axes = FALSE,
    cex.lab = 1.5,
    xlim = c(min(Wavelength), max(Wavelength))
  )
  
  axis(1,
       at = seq(min(Wavelength), max(Wavelength), by = 50),
       labels = seq(min(Wavelength), max(Wavelength), by = 50),
       cex.axis = 1.25, lwd = 3)
  
  box(lwd = 3)
  
  # ---- Overlay class-average spectra ----
  lines(avg_spec$Wavelength, avg_spec$China,  col = "#E74C3C",    lwd = 3)
  lines(avg_spec$Wavelength, avg_spec$Japan,  col = "gold",       lwd = 3)
  lines(avg_spec$Wavelength, avg_spec$Korea,  col = "dodgerblue", lwd = 3)
}

# Example:
# plot_mean_shap()

###############################################################################
cat("\nSHAP computation completed successfully.\n")
###############################################################################
