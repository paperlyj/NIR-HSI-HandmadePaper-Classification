# 📘 NIR Hyperspectral Imaging Dataset for Traditional Handmade Papers

This repository contains the dataset and analysis code used in the study:

**“Classification of traditional handmade papers using NIR hyperspectral imaging”**  
Yong Ju Lee et al., submitted to *Royal Society Open Science*.

---

## 📁 Contents

- NIR spectral dataset (1250–1700 nm)  
- Code for preprocessing, PCA, k-NN, SVM, and ANN classification  
- SHAP interpretation scripts (feature attribution for ANN)

---

## 📄 Data Description

- **260 NIR spectra** from 26 traditional handmade papers  
  - Hanji (Korea)  
  - Washi (Japan)  
  - Xuan (China)

- **Instrument**  
  - Resonon Pika NIR-320 (900–1700 nm)

- **Preprocessing**  
  - Savitzky–Golay 2nd derivative (m = 2, p = 3, w = 7)  
  - L2 (Euclidean) normalization

---

## 📦 Files

### Dataset
- `NIR_Summary.xlsx`  
  - Sheet **CL_Country**: country-level labels  
  - Sheet **CL_Product**: product-level labels

### Analysis Scripts
- `Principal component analysis (PCA).R`  
- `k nearest neighbor (KNN).R`  
- `Support vector machine (SVM).R`  
- `Feedforward neural network (FNN).R`  
- `SHAP Interpretation for FNN.R`

Each script performs complete preprocessing, training, validation, and visualization.

---

## 🔧 Requirements

Analyses were performed in **R (version 4.4.2)** using:

- `openxlsx`  
- `prospectr`  
- `caret`  
- `kknn`  
- `kernlab`  
- `ANN2`  
- `fastshap`  
- `ggplot2`

Record exact versions with:

```r
sessionInfo()
