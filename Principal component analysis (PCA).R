###############################################################################
# NIR HSI PCA Analysis for Traditional Handmade Papers
# - Savitzky–Golay 2nd derivative + L2 normalization PCA
# - Raw spectrum + L2 normalization PCA
# - PCA loading plot
#
# Input  : NIR_Summary.xlsx (sheet = "CL_Product")
# Output : PCA score plots and loading plot
###############################################################################

## 0. Load Required Libraries --------------------------------------------------

library(openxlsx)
library(prospectr)   # Savitzky-Golay filter
library(dplyr)
library(ggplot2)

## 1. Import Data --------------------------------------------------------------

# Path to your data file (GitHub users will place this under /data/)
data_file <- "NIR_Summary.xlsx"   # e.g., data/NIR_Summary.xlsx

# Read the dataset
nir_raw <- read.xlsx(xlsxFile = data_file, sheet = "NIR_Summary")

# Extract sample names
Sample <- nir_raw$Sample
nir_raw$Sample <- as.factor(nir_raw$Sample)

# Wavelength labels
Wavelength <- colnames(nir_raw[, -1])

## 2. L2 Normalization Function ------------------------------------------------

euclidean_normalize <- function(x) {
  norm <- sqrt(sum(x^2))
  x / norm
}

###############################################################################
# 2ND DERIVATIVE + L2 NORMALIZATION PCA
###############################################################################

## 3. Apply Savitzky–Golay 2nd Derivative -------------------------------------

sg_2d <- apply(nir_raw[, -1], 1,
               function(x) savitzkyGolay(x, m = 2, p = 3, w = 7))

data_2d <- data.frame(Sample = Sample, t(sg_2d))

# Clean wavelength labels
Wavenumber_2d <- sub("^X", "", colnames(data_2d[, -1]))

## 4. L2 Normalize the Derivative Spectrum ------------------------------------

norm_2d <- apply(data_2d[, -1], 1, euclidean_normalize)
data_2d <- data.frame(Sample = Sample, t(norm_2d))

## 5. PCA ----------------------------------------------------------------------

pr_2d <- prcomp(data_2d[, -1], scale. = FALSE, rank. = 7)
summary(pr_2d)

## 6. Add Metadata -------------------------------------------------------------

df_2d <- data_2d %>%
  mutate(
    Country = case_when(
      grepl("^China", Sample) ~ "China",
      grepl("^Japan", Sample) ~ "Japan",
      TRUE                    ~ "Korea"
    ),
    Product = gsub(".*No\\.\\s*", "", Sample) |> factor()
  )

n_prod <- nlevels(df_2d$Product)
shape_codes <- 0:(n_prod - 1)   # pch values

## 7. PCA Score Plot (PC1 vs PC2) ---------------------------------------------

ggplot(df_2d,
       aes(x = pr_2d$x[, 1], y = pr_2d$x[, 2],
           color = Country, shape = Product, fill = Country)) +
  geom_point(size = 6.5, stroke = 1.1) +
  scale_color_manual(values = c(China="#E74C3C", Korea="#2E8B57", Japan="#1E90FF")) +
  scale_fill_manual(values = c(China="#F5B7B1", Korea="#ABEBC6", Japan="#AED6F1")) +
  scale_shape_manual(values = shape_codes) +
  guides(color = guide_legend(title = "Country"),
         shape = guide_legend(title = "Product No.", ncol = 3),
         fill = "none") +
  theme_classic(base_size = 12) +
  theme(
    legend.position = "none",
    axis.title      = element_text(size = 16, face = "bold"),
    axis.text       = element_text(size = 14),
    panel.border    = element_rect(color = "black", fill = NA, linewidth = 1.2)
  ) +
  xlab(expression(bold("PC1 (55.0%)"))) +
  ylab(expression(bold("PC2 (21.6%)"))) +
  coord_cartesian(xlim = c(-0.2, 0.4), ylim = c(-0.3, 0.2)) +
  scale_x_continuous(breaks = seq(-0.3, 0.4, by = 0.1),
                     labels = scales::number_format(accuracy = 0.1)) +
  scale_y_continuous(breaks = seq(-0.3, 0.2, by = 0.1))

###############################################################################
# RAW SPECTRUM + L2 NORMALIZATION PCA
###############################################################################

## 8. Normalize Raw Spectra ----------------------------------------------------

norm_raw <- apply(nir_raw[, -1], 1, euclidean_normalize)
data_raw_norm <- data.frame(Sample = Sample, t(norm_raw))

## 9. PCA ----------------------------------------------------------------------

pr_raw <- prcomp(data_raw_norm[, -1], scale. = FALSE, rank. = 7)
summary(pr_raw)

## 10. Metadata ----------------------------------------------------------------

df_raw <- data_raw_norm %>%
  mutate(
    Country = case_when(
      grepl("^China", Sample) ~ "China",
      grepl("^Japan", Sample) ~ "Japan",
      TRUE ~ "Korea"
    ),
    Product = gsub(".*No\\.\\s*", "", Sample) |> factor()
  )

shape_codes <- 0:(nlevels(df_raw$Product) - 1)

## 11. PCA Score Plot (Raw + L2, PC1 vs PC2) ----------------------------------

ggplot(df_raw,
       aes(x = pr_raw$x[, 1], y = pr_raw$x[, 2],
           color = Country, shape = Product, fill = Country)) +
  geom_point(size = 6.5, stroke = 1.1) +
  scale_color_manual(values = c(China="#E74C3C", Korea="#2E8B57", Japan="#1E90FF")) +
  scale_fill_manual(values = c(China="#F5B7B1", Korea="#ABEBC6", Japan="#AED6F1")) +
  scale_shape_manual(values = shape_codes) +
  theme_classic(base_size = 12) +
  theme(
    legend.position = "none",
    axis.title = element_text(size = 16, face = "bold")
  ) +
  xlab(expression(bold("PC1 (84.9%)"))) +
  ylab(expression(bold("PC2 (9.8%)"))) +
  coord_cartesian(xlim = c(-0.02, 0.03), ylim = c(-0.008, 0.01)) +
  scale_x_continuous(breaks = seq(-0.02, 0.03, by = 0.01)) +
  scale_y_continuous(breaks = seq(-0.01, 0.01, by = 0.002))

###############################################################################
# PCA LOADING PLOT
###############################################################################

## 12. Loading Plot for PC1 and PC2 ------------------------------------------

wavelength_seq <- seq(1250, 1700, length.out = ncol(pr_2d$rotation))

plot(wavelength_seq, pr_2d$rotation[, 1],
     type = "l", lwd = 3, col = "red",
     xlab = "Wavelength (nm)", ylab = "PCA Loading",
     xaxs = "i", yaxs = "i",
     ylim = c(-0.5, 0.4))

lines(wavelength_seq, pr_2d$rotation[, 2],
      type = "l", lwd = 3, col = "black")

legend("bottomleft",
       legend = c("PC1 (55.0%)", "PC2 (21.6%)"),
       col = c("red", "black"), lwd = 3, bty = "n")
