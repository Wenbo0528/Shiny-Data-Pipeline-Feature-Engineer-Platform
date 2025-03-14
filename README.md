# Shiny-Data-Pipeline-and-Feature-Engineering-Platform

## Project Overview
This project is a comprehensive web application built with **R Shiny**. It streamlines the data workflow by integrating **data upload**, **cleaning & preprocessing**, **feature engineering**, **exploratory data analysis (EDA)**, and **downloading** functions into one interactive platform. Users can either upload local files or use built-in R datasets.

---

## Code & Tools Used
- **Programming Language**: R  
- **Core Shiny App**: [data_explorer_app_final.R](./data_explorer_app_final.R)

## Instructions
1. **Clone or download** this repository to your local machine.
2. **Install** the required packages listed above.
3. Open the **data_explorer_app_final.R** file in RStudio (or any R environment).
4. Click **Run App**. A Shiny app will open in your browser.
5. Use the **tabs** to upload, clean/preprocess, engineer features, explore, and download your processed data.
6. Select **Reset App** when you need to start fresh.

---

## Main Functional Tabs
1. **User Guide**  
   - Introductory instructions on navigating each tab in the Shiny app.

2. **Data Upload**  
   - Upload files (`.csv`, `.xlsx`, `.json`, `.rds`) or select built-in datasets (e.g., `mtcars`).
   - Preview loaded datasets.

3. **Data Cleaning & Preprocessing**  
   - Remove duplicates, handle missing values (remove or impute), and detect outliers (Z-score/IQR).  
   - Normalize/standardize numeric columns, encode categorical variables, and convert numeric-like data.  
   - Optionally remove selected columns and round numeric values.

4. **Feature Engineering**  
   - Apply transformations (log, square root, square, Box-Cox, power).  
   - Create additional features (time-based, text-based, statistical summaries).  
   - Conduct feature selection (variance threshold, correlation, Lasso) and dimensionality reduction (PCA, LDA).  
   - Remove or keep specific columns post-transformation.

5. **Exploratory Data Analysis (EDA)**  
   - Generate histograms, bar charts, boxplots, scatter plots, and correlation heatmaps.  
   - Access interactive plots (via **plotly**) and a statistical summary of numeric columns.

6. **Download & Reset**  
   - Download the processed dataset as `.xlsx`, `.csv`, `.json`, or `.rds`.  
   - Reset the application to clear all uploaded and processed data.

---

**Required R Packages** (install them before running the app):
- shiny  
- shinyjs  
- shinythemes  
- DT  
- dplyr  
- openxlsx  
- stringr  
- readr  
- readxl  
- jsonlite  
- ggplot2  
- plotly  
- lubridate  
- ggcorrplot  
- recipes  
- caret  
- reshape2  
- MASS  
- rmarkdown  
- DBI  
- shinyWidgets  

(For **KNN imputation**, install `DMwR`; for **Lasso** selection, install `glmnet`.)

---

## Collaborators
- Wenbo Liu ( Wenbo0528 )  
- Fatih Euysal ( fatihEuysal )  
- Xiaoying Wang ( XiaoyingWang412 )  
- Julieta Caroppo ( julieta87 )

---

## License
This project is for educational purposes and is licensed under the **MIT License**.
