# Load necessary libraries
library(shiny)
library(shinyjs)
library(shinythemes)
library(DT)
library(dplyr)
library(openxlsx)
library(stringr)
library(readr)     # For CSV
library(readxl)    # For Excel
library(jsonlite)  # For JSON
library(ggplot2)
library(plotly)    # For interactive plots
library(lubridate)
library(ggcorrplot)
library(recipes)
library(caret)
library(reshape2)
library(MASS)           # For Box-Cox
library(rmarkdown)      # For report generation
library(DBI)            # For database connections
library(shinyWidgets)
library(rsconnect)

# 1. Define UI
ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      body { background: #f8f8f8; color: #333; overflow-y: auto; }
      .container { background: #ffffff; padding: 20px; border-radius: 8px; margin-top: 20px; margin-bottom: 20px; }
      .navbar, .navbar-default { background-color: #222222 !important; border-color: #222222; }
      .navbar .navbar-nav > li > a { color: #ffffff !important; }
      .dataTables_wrapper .dataTables_paginate .paginate_button { color: #333333 !important; }
      /* Tooltip styling for datatable headers */
      th { position: relative; }
      th:hover::after {
        content: attr(title);
        position: absolute; top: 100%; left: 0;
        background: #333; color: #fff; padding: 3px 5px;
        white-space: nowrap; z-index: 100; border-radius: 3px;
      }
      /* Custom styles for selectize multi-select */
      .selectize-control.multi .selectize-input [data-value] {
        padding: 2px 20px 2px 8px !important; position: relative;
      }
      .selectize-control.multi .selectize-input [data-value] .remove {
        position: absolute; top: 0; right: 0; width: 20px; height: 100%;
        display: flex; align-items: center; justify-content: center;
        border-left: 1px solid rgba(0,0,0,0.1); background: rgba(0,0,0,0.05); cursor: pointer;
      }
      .selectize-control.multi .selectize-input [data-value] .remove:hover { background: rgba(0,0,0,0.1); }
    "))
  ),
  theme = shinytheme("flatly"),  # Use a modern shiny theme
  useShinyjs(),                  # Enable shinyjs for extra interactivity
  titlePanel("Data Explorer & Comparison Tool"),
  navbarPage("",
             
             # 1.1 User Guide Part
             tabPanel("User Guide", fluidRow(
               column(12,
                      h3("Overview"),
                      p("This web application is designed to provide an interactive and user-friendly platform for data analysis, enabling users to seamlessly upload, clean, preprocess, engineer features, and explore datasets. The app simplifies data handling and visualization for data scientists and analysts."),
                      h4("Key Features:"),
                      tags$ul(
                        tags$li(strong("Data Upload:"), "Supports CSV, Excel, JSON, and RDS files"),
                        tags$li(strong("Data Cleaning:"), "Remove duplicates, handle missing values, outliers, normalize/standardize, encode categorical variables, and convert numeric-like columns"),
                        tags$li(strong("Feature Engineering:"), "Apply mathematical transformations, create custom columns, and rename columns"),
                        tags$li(strong("EDA:"), "Generate interactive histograms, boxplots, bar charts, scatter plots, and correlation heatmaps"),
                        tags$li(strong("Download:"), "Export the cleaned data"),
                        tags$li("Responsive and user-friendly interface for smooth data exploration")
                      ),
                      h4("Navigation:"),
                      p("Use the tabs to navigate between Data Upload, Data Cleaning, Feature Engineering, EDA, and Download sections.")
               )
             ))
             ,
             
             # 1.2 Data Upload part
             tabPanel("Data Upload", sidebarLayout(
               sidebarPanel(
                 h4(strong("Data Source Options")),
                 radioButtons("dataSource", "Data Source:",
                              choices = c("Upload File(s)" = "upload", "Use Built-in Dataset(s)" = "builtin"),
                              selected = "upload"),
                 conditionalPanel(
                   condition = "input.dataSource == 'upload'",
                   fileInput("file", "Choose dataset file(s) (Supports: .csv, .xlsx, .json, .rds):",
                             accept = c(".csv", ".xlsx", ".json", ".rds"),
                             multiple = TRUE)
                 ),
                 conditionalPanel(
                   condition = "input.dataSource == 'builtin'",
                   selectizeInput("builtinDataset", "Select built-in dataset(s):",
                                  choices = c("mtcars", "iris", "ToothGrowth"),
                                  multiple = TRUE,
                                  selected = "mtcars",
                                  options = list(plugins = list('remove_button')))
                 ),
                 actionButton("loadData", "Load Data")
               ),
               mainPanel(
                 div(class = "container",
                     h4("Preview Loaded Dataset"),
                     selectInput("uploadPreview", "Select Dataset to Preview:", choices = NULL),
                     DTOutput("dataPreview", height = "300px")
                 )
               )
             )),
             
             # 1.3 Data Cleaning & Preprocessing Part
             tabPanel("Data Cleaning & Preprocessing",
                      sidebarLayout(
                        sidebarPanel(
                          h4(strong("Cleaning Options")),
                          
                          # 1. Dataset Selection
                          selectInput("activeDataset", "Select Active Dataset:", choices = NULL),
                          
                          # 2. Data Handling
                          strong("General Data Handling"),
                          checkboxInput("removeDup", "Remove duplicate rows", value = FALSE),
                          checkboxInput("handleMissing", "Handle Missing Values", value = FALSE),
                          conditionalPanel(
                            condition = "input.handleMissing == true",
                            radioButtons("missOpt", "Missing Value Handling Method:",
                                         choices = c("Remove Rows" = "remove", 
                                                     "Impute with Mean" = "mean",
                                                     "Impute with Median" = "median",
                                                     "Impute with Mode" = "mode",
                                                     "Impute with KNN" = "knn"),
                                         selected = "remove")
                          ),
                          checkboxInput("handleOutliers", "Handle Outliers", value = FALSE),
                          conditionalPanel(
                            condition = "input.handleOutliers == true",
                            radioButtons("outlierMethod", "Outlier Detection Method:",
                                         choices = c("Z-score" = "zscore", "IQR" = "iqr"),
                                         selected = "zscore")
                          ),
                          checkboxInput("roundValues", "Round numeric values", value = FALSE),
                          conditionalPanel(
                            condition = "input.roundValues == true",
                            numericInput("roundDigits", "Round to how many decimal places:", value = 2, min = 0, max = 10)
                          ),
                          
                          # 3. Feature Transformations
                          strong("Feature Transformations"),
                          checkboxInput("normalizeData", "Normalize Numerical Features", value = FALSE),
                          checkboxInput("standardizeData", "Standardize Numerical Features", value = FALSE),
                          checkboxInput("encodeCategorical", "Encode Categorical Variables", value = FALSE),
                          checkboxInput("convertNumeric", "Convert numeric-like character columns to numeric", value = FALSE),
                          
                          # 4. Column Management
                          strong("Column Management (After doing previous step)"),
                          checkboxInput("removeColumns", "Select Columns to Remove", value = FALSE),
                          conditionalPanel(
                            condition = "input.removeColumns == true",
                            checkboxGroupInput("colsToRemove", "Columns to Remove:", choices = NULL)
                          ),
                          br(),
                          
                          # 5. Execute Cleaning
                          actionButton("clean", "Clean Data")
                        ),
                        
                        mainPanel(
                          div(class = "container",
                              h4("Cleaning Summary"),
                              uiOutput("cleaningSummary"),
                          ),
                          div(class = "container",
                              h4("Before & After Data Cleaning"),
                              tabsetPanel(
                                tabPanel("Before Cleaning", DTOutput("beforeCleanPreview", height = "350px")),
                                tabPanel("After Cleaning", DTOutput("cleanPreview", height = "350px"))
                              )
                          )
                        )
                      )
             )
             ,
             
             # 1.4 Feature Engineering Part
             tabPanel("Feature Engineering",
                      sidebarLayout(
                        sidebarPanel(
                          h4(strong("Feature Engineering Options")),
                          
                          # Select Active Dataset
                          selectInput("featureActiveDataset", "Select Active Dataset:", choices = NULL),
                          
                          # Mathematical transformations
                          checkboxGroupInput("mathTransform", "Select Transformations:",
                                             choices = c("Log" = "log", "Square Root" = "sqrt", "Square" = "square", "Box-Cox" = "boxcox", "Power" = "power")),
                          
                          # Feature selection
                          checkboxGroupInput("featureSelection", "Feature Selection Methods:",
                                             choices = c("Variance Threshold (0.01)" = "var_thresh", 
                                                         "Correlation Analysis (>0.8)" = "corr",
                                                         "Lasso Regularization" = "lasso")),
                          
                          # Dimensionality reduction
                          checkboxGroupInput("dimReduction", "Dimensionality Reduction:",
                                             choices = c("PCA (2 dimensions)" = "pca", "LDA (2 dimensions)" = "lda")),
                          
                          # Additional feature engineering
                          checkboxGroupInput("additionalFeatureEng", "Additional Feature Engineering:",
                                             choices = c("Extract Time Features" = "time",
                                                         "Extract Text Features" = "text",
                                                         "Generate Statistical Features" = "stats")),
                          strong("Column Management (After doing previous step)"),
                          checkboxInput("removeColumnsF", "Select Columns to Remove", value = FALSE),
                          conditionalPanel(
                            condition = "input.removeColumnsF == true",
                            checkboxGroupInput("colsToRemoveF", "Columns to Remove:", choices = NULL)
                          ),

                          
                          br(),
                          actionButton("applyFeatureEng", "Apply Feature Engineering")
                        ),
                        
                        # Right-side main panel
                        mainPanel(
                          div(class = "container",
                              h4("Feature Engineering Summary"),
                              uiOutput("featureEngineeringSummary")
                          ),
                          div(class = "container",
                              h4("Before & After Feature Engineering"),
                              tabsetPanel(
                                tabPanel("Before Feature Engineering", DTOutput("beforeFeatureEngPreview", height = "350px")),
                                tabPanel("After Feature Engineering", DTOutput("featureEngPreview", height = "350px"))
                              )
                          )
                        )
                      ))
             ,
             
             # 1.5 Exploratory Data Analysis Part
             tabPanel("Exploratory Data Analysis", 
                      sidebarLayout(
                        sidebarPanel(
                          h4(strong("Exploratory Data Analysis Options")),
                          
                          # Select Active Dataset
                          selectInput("edaActiveDataset", "Select Active Dataset:", choices = NULL),
                          
                          # Select Chart Type First
                          radioButtons("plotType", "Select Chart Type:", 
                                       choices = c("Histogram" = "hist", "Bar Chart" = "bar", 
                                                   "Boxplot" = "boxplot", "Scatter Plot" = "scatter", 
                                                   "Correlation Heatmap" = "heatmap"),
                                       selected = "hist"),
                          
                          # Dynamic X and Y axis selection based on chart type
                          conditionalPanel(
                            condition = "input.plotType != 'heatmap'", 
                            selectInput("xVar", "Select X-axis Variable:", choices = NULL)
                          ),
                          
                          conditionalPanel(
                            condition = "input.plotType == 'scatter' || input.plotType == 'boxplot'", 
                            selectInput("yVar", "Select Y-axis Variable:", choices = NULL, selected = NULL)
                          ),
                          
                          sliderInput("alpha", "Opacity:", min = 0.1, max = 1, value = 0.7),
                          actionButton("plotData", "Generate Plot", class = "btn-primary")
                        ),
                        
                        # Main panel with separate sections for visualization and statistical summary
                        mainPanel(
                          div(class = "container",
                              h4("Data Exploration Visualization"),
                              conditionalPanel(
                                condition = "input.plotType == 'heatmap'", 
                                plotOutput("heatmapPlot", height = "500px")
                              ),
                              conditionalPanel(
                                condition = "input.plotType != 'heatmap'", 
                                plotlyOutput("edaPlot", height = "500px")
                              )
                          ),
                          
                          br(),
                          
                          div(class = "container",
                              h4("Statistical Summary"),
                              uiOutput("statSummary")
                          )
                        )
                      )
             )
             ,
             
             # 1.6 Download & Reset Tab
             tabPanel("Download & Reset",
                      sidebarLayout(
                        sidebarPanel(
                          h4(strong("Download Processed Data")),
                          
                          # Select dataset
                          selectInput("activeDatasetDL", "Select Dataset to Download:", choices = NULL),
                          
                          # Select file format
                          selectInput("downloadFormat", "Select File Format:",
                                      choices = c("Excel (.xlsx)" = "xlsx",
                                                  "CSV (.csv)" = "csv",
                                                  "JSON (.json)" = "json",
                                                  "RDS (.rds)" = "rds")),
                          
                          # Download button
                          downloadButton("downloadData", "Download Processed Data"),
                          
                          # Reset button
                          hr(),  # Add a horizontal line for separation
                          actionButton("reset", "Reset App", class = "btn-danger")
                        ),
                        
                        mainPanel(
                          h4("Instructions"),
                          p("Select a dataset and file format to download the processed data."),
                          p("Click 'Reset App' to clear all uploaded datasets and start fresh.")
                        )
                      )
             )
             
             
  )
)


# 2. Define Server logic
server <- function(input, output, session) {
  rv <- reactiveValues(data = list(), cleaned = list(), featured = list())
  
  # 2.1 Data Load Server
  observeEvent(input$loadData, {
    req(input$dataSource)
    
    # 1) If the data source is an uploaded file
    if (input$dataSource == "upload") {
      req(input$file)
      data_list <- list()
      
      # Attempt to parse each uploaded file
      for (i in seq_len(nrow(input$file))) {
        fname <- input$file$name[i]
        ext   <- tolower(tools::file_ext(fname))
        datap <- input$file$datapath[i]
        data  <- NULL
        
        # Handle different file types based on extensions
        if (ext == "csv") {
          data <- read_csv(datap)
          
        } else if (ext %in% c("xlsx", "xls")) {
          data <- read_excel(datap)
          
        } else if (ext == "json") {
          # Try parsing with fromJSON()
          data <- tryCatch(
            fromJSON(datap, flatten = TRUE),
            error = function(e) {
              # If it fails, attempt line-by-line parsing using stream_in()
              showNotification(
                paste("Standard JSON parsing for", fname, "failed. Trying line-by-line parsing."),
                type = "warning"
              )
              con <- file(datap, open = "r")
              on.exit(close(con), add = TRUE)
              
              # Attempt parsing again
              json_data <- tryCatch(
                {
                  out <- stream_in(con, verbose = FALSE)
                  out
                },
                error = function(e2) {
                  showNotification(
                    paste("Failed to parse JSON file", fname, "using line-by-line. Invalid JSON format?"),
                    type = "error"
                  )
                  NULL
                }
              )
              json_data
            }
          )
          
        } else if (ext == "rds") {
          data <- tryCatch(
            readRDS(datap),
            error = function(e) {
              showNotification(paste("Invalid RDS file format for", fname), type = "error")
              NULL
            }
          )
        } else {
          showNotification(paste("Unsupported file type:", ext, "for file:", fname), type = "error")
          next  # Skip unsupported file types
        }
        
        # Convert parsed data to a data frame (if not NULL)
        if (!is.null(data) && !is.data.frame(data)) {
          data <- as.data.frame(data)
        }
        
        # If the data was successfully parsed
        if (!is.null(data)) {
          data_list[[fname]] <- data
          # Notify the user that the file was loaded successfully
          showNotification(
            paste("File", fname, "loaded successfully!"),
            type = "message"
          )
        }
      }
      
      # Store the successfully parsed datasets
      rv$dataList <- data_list
      
      # If no dataset was successfully loaded, notify the user
      if (length(data_list) == 0) {
        showNotification("No valid dataset has been loaded. Please check your files.", type = "error")
      }
      
      # 2) If the data source is a built-in dataset
    } else {
      req(input$builtinDataset)
      data_list <- setNames(lapply(input$builtinDataset, get), input$builtinDataset)
      rv$dataList <- data_list
      showNotification("Built-in dataset(s) loaded successfully!", type = "message")
    }
    
    # If rv$dataList is not empty, update the dropdown menu to allow dataset selection for preview
    if (length(rv$dataList) > 0) {
      updateSelectInput(session, "uploadPreview",
                        choices  = names(rv$dataList),
                        selected = names(rv$dataList)[1])
    }
  })
  
  # Display the uploaded or selected raw dataset
  output$dataPreview <- renderDT({
    req(rv$dataList, input$uploadPreview)
    datatable(rv$dataList[[input$uploadPreview]], options = list(scrollX = TRUE))
  })
  
  # Display the dataset structure
  output$dataStructure <- renderPrint({
    req(rv$dataList, input$uploadPreview)
    str(rv$dataList[[input$uploadPreview]])
  })
  
  # Display raw dataset in "Before Cleaning" section of the Data Cleaning page
  output$beforeCleanPreview <- renderDT({
    req(rv$dataList, input$uploadPreview)
    datatable(rv$dataList[[input$uploadPreview]], options = list(scrollX = TRUE))
  })

  
  
  

  # 2.2 Data Cleaning & Preprocessing Server
  # Monitor changes in `activeDataset`, update Before Cleaning preview, and hide After Cleaning
  observe({
    req(rv$dataList, input$activeDataset)
    
    dataset_name <- input$activeDataset
    
    # Before Cleaning: Immediately display the original dataset
    output$beforeCleanPreview <- renderDT({
      req(rv$dataList[[dataset_name]])
      datatable(rv$dataList[[dataset_name]], options = list(scrollX = TRUE))
    })
    
    # After Cleaning: Display an empty table until Clean Data is clicked
    output$cleanPreview <- renderDT({
      datatable(data.frame(), options = list(scrollX = TRUE))
    })
  })
  
  # Monitor Clean Data button click and update After Cleaning preview
  observeEvent(input$clean, {
    req(rv$dataList, input$activeDataset)
    
    dataset_name <- input$activeDataset
    data <- rv$dataList[[dataset_name]]  # Retrieve original dataset
    
    # Remove duplicates
    if (input$removeDup) {
      data <- distinct(data)
    }
    
    # Handle missing values
    if (input$handleMissing) {
      if (input$missOpt == "remove") {
        data <- na.omit(data)
      } else if (input$missOpt == "mean") {
        data <- data %>% mutate(across(where(is.numeric), ~ ifelse(is.na(.), mean(., na.rm = TRUE), .)))
      } else if (input$missOpt == "median") {
        data <- data %>% mutate(across(where(is.numeric), ~ ifelse(is.na(.), median(., na.rm = TRUE), .)))
      } else if (input$missOpt == "mode") {
        Mode <- function(x) { ux <- unique(x); ux[which.max(tabulate(match(x, ux)))] }
        data <- data %>% mutate(across(where(is.character), ~ ifelse(is.na(.), Mode(.), .)))
      } else if (input$missOpt == "knn") {
        if (!requireNamespace("DMwR", quietly = TRUE)) {
          showNotification("KNN Imputation requires 'DMwR' package. Please install it.", type = "error")
        } else {
          data <- DMwR::knnImputation(data)
        }
      }
    }
    
    # Handle outliers
    if (input$handleOutliers) {
      if (input$outlierMethod == "zscore") {
        data <- data %>% mutate(across(where(is.numeric), ~ case_when(
          abs((. - mean(., na.rm = TRUE)) / sd(., na.rm = TRUE)) > 3 ~ NA_real_,
          TRUE ~ .
        )))
      } else if (input$outlierMethod == "iqr") {
        data <- data %>% mutate(across(where(is.numeric), ~ {
          Q1 <- quantile(., 0.25, na.rm = TRUE)
          Q3 <- quantile(., 0.75, na.rm = TRUE)
          IQR_value <- Q3 - Q1
          case_when(
            . < (Q1 - 1.5 * IQR_value) ~ NA_real_,
            . > (Q3 + 1.5 * IQR_value) ~ NA_real_,
            TRUE ~ .
          )
        }))
      }
    }
    
    # Standardization & Normalization
    if (input$normalizeData) {
      data <- data %>% mutate(across(where(is.numeric), ~ (. - min(., na.rm = TRUE)) / (max(., na.rm = TRUE) - min(., na.rm = TRUE))))
    }
    if (input$standardizeData) {
      data <- data %>% mutate(across(where(is.numeric), ~ (. - mean(., na.rm = TRUE)) / sd(., na.rm = TRUE)))
    }
    
    # Categorical encoding
    if (input$encodeCategorical) {
      data <- data %>% mutate(across(where(is.character), as.factor))
    }
    
    # Numeric conversion
    if (input$convertNumeric) {
      data <- data %>% mutate(across(where(is.character), ~ ifelse(suppressWarnings(!is.na(as.numeric(.))), as.numeric(.), .)))
    }
    
    # Remove selected columns
    if (!is.null(input$colsToRemove) && length(input$colsToRemove) > 0) {
      remaining_cols <- setdiff(names(data), input$colsToRemove)
      
      # Prevent the user from removing all columns
      if (length(remaining_cols) > 0) {
        data <- data[, remaining_cols, drop = FALSE]
      } else {
        showNotification("Cannot remove all columns!", type = "error")
        return()
      }
    }
    
    # Rounding values
    if (input$roundValues) {
      data <- data %>% mutate(across(where(is.numeric), ~ round(., input$roundDigits)))
    }
    
    # Update cleaned data
    rv$cleaned[[dataset_name]] <- data
    
    # Update After Cleaning preview
    output$cleanPreview <- renderDT({
      req(rv$cleaned[[dataset_name]])
      datatable(rv$cleaned[[dataset_name]], options = list(scrollX = TRUE))
    })
    
    # Update Before Feature Engineering preview
    output$beforeFeatureEngPreview <- renderDT({
      req(rv$cleaned[[dataset_name]])
      datatable(rv$cleaned[[dataset_name]], options = list(scrollX = TRUE))
    })
    
    # Update `colsToRemove` checkbox group
    updateCheckboxGroupInput(session, "colsToRemove", choices = names(data), selected = NULL)
    
    # Update Cleaning Summary
    output$cleaningSummary <- renderUI({
      req(rv$dataList, rv$cleaned, input$activeDataset)
      
      original_rows <- nrow(rv$dataList[[dataset_name]])
      cleaned_rows <- nrow(rv$cleaned[[dataset_name]])
      removed_cols <- ifelse(length(input$colsToRemove) > 0, paste(input$colsToRemove, collapse = ", "), "None")
      
      tags$div(
        h5("Cleaning Summary:"),
        p(paste("Dataset:", dataset_name)),
        p(paste("Original Rows:", original_rows)),
        p(paste("Rows After Cleaning:", cleaned_rows)),
        p(paste("Columns Removed:", removed_cols))
      )
    })
    
    # Notify the user
    showNotification(paste("Data cleaning completed successfully for:", dataset_name), type = "message")
  })
  
  # Dynamically update `activeDataset` selection input
  observe({
    req(rv$dataList)
    updateSelectInput(session, "activeDataset", choices = names(rv$dataList), selected = names(rv$dataList)[1])
  })

  
  
  
  # 2.3 Feature Engineering Server
  # Monitor changes in `activeDataset`, update Before Feature Engineering preview, and hide After Feature Engineering
  observe({
    req(rv$cleaned, input$featureActiveDataset)
    
    dataset_name <- input$featureActiveDataset
    
    # Before Feature Engineering: Display the cleaned dataset
    output$beforeFeatureEngPreview <- renderDT({
      req(rv$cleaned[[dataset_name]])
      datatable(rv$cleaned[[dataset_name]], options = list(scrollX = TRUE))
    })
    
    # After Feature Engineering: Initialize as empty
    output$featureEngPreview <- renderDT({
      datatable(data.frame(), options = list(scrollX = TRUE))
    })
  })
  
  # Listen for the "Apply Feature Engineering" button
  observeEvent(input$applyFeatureEng, {
    req(rv$cleaned, input$featureActiveDataset)
    
    dataset_name <- input$featureActiveDataset
    data <- rv$cleaned[[dataset_name]]  # Retrieve the cleaned dataset
    
    # 1. Mathematical transformations
    if (!is.null(input$mathTransform)) {
      if ("log" %in% input$mathTransform) {
        data <- data %>% mutate(across(where(is.numeric), ~ log(. + 1), .names = "log_{.col}"))
      }
      if ("sqrt" %in% input$mathTransform) {
        data <- data %>% mutate(across(where(is.numeric), sqrt, .names = "sqrt_{.col}"))
      }
      if ("square" %in% input$mathTransform) {
        data <- data %>% mutate(across(where(is.numeric), ~ .^2, .names = "square_{.col}"))
      }
      
      # Box-Cox transformation
      if ("boxcox" %in% input$mathTransform) {
        lambda <- 0.5  # This can be extended to allow user input
        data <- data %>% mutate(across(where(is.numeric), ~ {
          x <- . + abs(min(., na.rm = TRUE)) + 1  # Ensure all values are positive
          if (lambda == 0) log(x) else (x^lambda - 1) / lambda
        }, .names = "boxcox_{.col}"))
      }
      
      # Power transformation
      if ("power" %in% input$mathTransform) {
        lambda <- 1.5  # This can be extended to allow user input
        data <- data %>% mutate(across(where(is.numeric), ~ .^lambda, .names = "power_{.col}"))
      }
    }
    
    # 2. Additional feature engineering
    if (!is.null(input$additionalFeatureEng)) {
      
      # Extract time features
      if ("time" %in% input$additionalFeatureEng) {
        date_cols <- names(data)[sapply(data, lubridate::is.Date)]
        
        if (length(date_cols) > 0) {
          data <- data %>%
            mutate(across(all_of(date_cols), list(
              year = ~ lubridate::year(.),
              month = ~ lubridate::month(.),
              day = ~ lubridate::day(.),
              weekday = ~ lubridate::wday(., label = TRUE)
            ), .names = "{.col}_{.fn}"))
        } else {
          showNotification("No date-type columns found for Extract Time Features.", type = "warning")
        }
      }
      
      # Extract text features
      if ("text" %in% input$additionalFeatureEng) {
        text_cols <- names(data)[sapply(data, is.character)]
        
        if (length(text_cols) > 0) {
          data <- data %>%
            mutate(across(all_of(text_cols), list(
              word_count = ~ str_count(., "\\S+"),
              char_count = ~ nchar(.)
            ), .names = "{.col}_{.fn}"))
        } else {
          showNotification("No text-type columns found for Extract Text Features.", type = "warning")
        }
      }
      
      # Generate statistical features
      if ("stats" %in% input$additionalFeatureEng) {
        num_cols <- names(data)[sapply(data, is.numeric)]
        
        if (length(num_cols) > 0) {
          data <- data %>%
            mutate(across(all_of(num_cols), list(
              mean = ~ mean(., na.rm = TRUE),
              var = ~ var(., na.rm = TRUE),
              median = ~ median(., na.rm = TRUE),
              min = ~ min(., na.rm = TRUE),
              max = ~ max(., na.rm = TRUE)
            ), .names = "{.col}_{.fn}"))
        } else {
          showNotification("No numeric columns found for Generate Statistical Features.", type = "warning")
        }
      }
    }

    # 3. Feature Selection
    if (!is.null(input$featureSelection)) {
      
      # Low variance feature selection
      if ("var_thresh" %in% input$featureSelection) {
        numeric_features <- data %>% dplyr::select_if(is.numeric)  # Previous code: select(data, where(is.numeric))
        
        if (ncol(numeric_features) > 0) {
          variances <- apply(numeric_features, 2, var, na.rm = TRUE)
          low_var_cols <- names(variances[variances <= 0.01])  # Set low variance threshold to 0.01
          if (length(low_var_cols) > 0) {
            data <- data %>% dplyr::select(-all_of(low_var_cols))  # Remove low variance features
            showNotification(paste("Removed low variance features:", paste(low_var_cols, collapse = ", ")), type = "message")
          }
        }
      }
      
      # Correlation analysis
      if ("corr" %in% input$featureSelection) {
        numeric_features <- data %>% dplyr::select_if(is.numeric)  # Previous code: select(data, where(is.numeric))
        
        if (ncol(numeric_features) > 1) {
          corr_matrix <- cor(numeric_features, use = "pairwise.complete.obs")
          high_corr <- caret::findCorrelation(corr_matrix, cutoff = 0.8)  # Set correlation threshold to 0.8
          if (length(high_corr) > 0) {
            removed_corr_features <- colnames(numeric_features)[high_corr]
            data <- data[, -high_corr, drop = FALSE]  # Remove highly correlated features
            showNotification(paste("Removed highly correlated features:", paste(removed_corr_features, collapse = ", ")), type = "message")
          }
        } else {
          showNotification("Correlation Analysis requires at least two numeric features.", type = "warning")
        }
      }
      
      # Lasso logistic regression feature selection (defaults to using the last categorical variable as the target variable)
      if ("lasso" %in% input$featureSelection) {
        categorical_vars <- names(data %>% dplyr::select_if(is.factor))
        numeric_features <- data %>% dplyr::select_if(is.numeric)
        
        if (length(categorical_vars) > 0 && ncol(numeric_features) > 1) {
          target_var <- tail(categorical_vars, 1)  # Default to the last categorical variable
          x <- as.matrix(numeric_features)
          y <- as.factor(data[[target_var]])
          
          # Select Lasso model family: use "binomial" for binary classification, "multinomial" for multi-class classification
          lasso_family <- ifelse(length(unique(y)) > 2, "multinomial", "binomial")
          
          if (requireNamespace("glmnet", quietly = TRUE)) {
            lasso_model <- glmnet::cv.glmnet(x, y, family = lasso_family, alpha = 1)
            selected_features <- coef(lasso_model, s = "lambda.min")  # Get Lasso-selected features
            
            # Process Lasso output
            if (lasso_family == "multinomial") {
              selected_features <- selected_features[[1]][-1]  # Retrieve Lasso-selected features for all classes
            } else {
              selected_features <- selected_features[-1]  # Keep only numerical features
            }
            
            selected_features <- names(selected_features[selected_features != 0])  # Select non-zero weight features
            if (length(selected_features) > 0) {
              data <- data %>% dplyr::select(all_of(selected_features), all_of(target_var))  # Keep only Lasso-selected features
              showNotification(paste("Lasso selected features:", paste(selected_features, collapse = ", ")), type = "message")
            } else {
              showNotification("Lasso did not select any features.", type = "warning")
            }
          } else {
            showNotification("Lasso feature selection requires 'glmnet' package.", type = "error")
          }
        } else {
          showNotification("Lasso requires a categorical target variable and at least two numeric features.", type = "warning")
        }
      }
    }
    
    
    # 4. Dimensionality Reduction
    if (!is.null(input$dimReduction)) {
      if ("pca" %in% input$dimReduction) {
        numeric_data <- data %>% select_if(is.numeric)
        
        if (ncol(numeric_data) >= 2) {
          pca_result <- prcomp(numeric_data, scale. = TRUE)
          pca_df <- as.data.frame(pca_result$x[, 1:2])  # Retain only the first two principal components
          colnames(pca_df) <- c("PCA_1", "PCA_2")
          data <- cbind(data, pca_df)
        } else {
          showNotification("PCA requires at least two numeric features and cannot be performed.", type = "error")
        }
        
      }
    }
    
    if ("lda" %in% input$dimReduction) {
      categorical_vars <- names(data %>% select_if(is.factor))
      numeric_data <- data %>% select_if(is.numeric)
      
      if (length(categorical_vars) > 0) {
        target_var <- tail(categorical_vars, 1)  # Default to the last categorical variable
        
        if (!is.null(input$ldaTargetVar) && input$ldaTargetVar %in% categorical_vars) {
          target_var <- input$ldaTargetVar  # Use user-selected target variable if provided
        }
        
        if (ncol(numeric_data) >= 2) {
          lda_result <- MASS::lda(numeric_data, grouping = data[[target_var]])
          lda_df <- as.data.frame(predict(lda_result)$x)
          colnames(lda_df) <- paste0("LDA_", seq_len(ncol(lda_df)))
          data <- cbind(data, lda_df)
        } else {
          showNotification("LDA requires at least two numeric features and cannot be performed.", type = "error")
        }
      } else {
        showNotification("LDA requires at least one categorical target variable, but none were found in the dataset.", type = "error")
      }
    }
    
    # Remove selected columns
    if (!is.null(input$colsToRemoveF) && length(input$colsToRemoveF) > 0) {
      remaining_cols <- setdiff(names(data), input$colsToRemoveF)
      
      if (length(remaining_cols) > 0) {
        data <- data[, remaining_cols, drop = FALSE]  # Keep only non-removed columns
      } else {
        showNotification("Cannot remove all columns!", type = "error")
        return()
      }
    }
    
    # 5. Update Feature Engineering results
    rv$featured[[dataset_name]] <- data
    
    # 6. Update After Feature Engineering preview
    output$featureEngPreview <- renderDT({
      req(rv$featured[[dataset_name]])
      datatable(rv$featured[[dataset_name]], options = list(scrollX = TRUE))
    })
    
    # 7. Update `colsToRemoveF` checkbox
    updateCheckboxGroupInput(session, "colsToRemoveF", choices = names(data), selected = NULL)
    
    # 8. Update Feature Engineering Summary
    output$featureEngineeringSummary <- renderUI({
      req(rv$cleaned, rv$featured, input$featureActiveDataset)
      
      original_cols <- names(rv$cleaned[[dataset_name]])
      engineered_cols <- names(rv$featured[[dataset_name]])
      added_cols <- setdiff(engineered_cols, original_cols)
      
      tags$div(
        h5("Feature Engineering Summary:"),
        p(paste("Dataset:", dataset_name)),
        p(paste("Original Columns:", length(original_cols))),
        p(paste("New Columns Added:", ifelse(length(added_cols) > 0, paste(added_cols, collapse = ", "), "None")))
      )
    })
    
    # 9. Notify user
    showNotification(paste("Feature Engineering applied to:", dataset_name), type = "message")
    
  })
  
  # Monitor `activeDataset` and update the selection dropdown
  observe({
    req(rv$cleaned)
    updateSelectInput(session, "featureActiveDataset", choices = names(rv$cleaned), selected = names(rv$cleaned)[1])
  })
  

  
  # 2.4 Exploratory Data Analysis Server
  # Monitor changes in `activeDataset` and update variable selection
  observe({
    req(rv$featured, input$edaActiveDataset)
    
    dataset_name <- input$edaActiveDataset
    current_data <- rv$featured[[dataset_name]]
    
    # Update X and Y variable selection
    updateSelectInput(session, "xVar", choices = names(current_data), selected = names(current_data)[1])
    updateSelectInput(session, "yVar", choices = c("", names(current_data)), selected = "")
  })
  
  # Statistical Summary Section
  output$statSummary <- renderUI({
    req(rv$featured, input$edaActiveDataset)
    
    dataset_name <- input$edaActiveDataset
    data <- rv$featured[[dataset_name]]
    
    # Select numeric columns
    numeric_cols <- sapply(data, is.numeric)
    numeric_data <- data[, numeric_cols, drop = FALSE]
    
    # If no numeric columns exist, display a warning message
    if (ncol(numeric_data) == 0) {
      return(tags$div(
        class = "alert alert-warning",
        "No numeric variables found in the dataset."
      ))
    }
    
    # Compute statistical summary
    summary_df <- data.frame(
      Variable = colnames(numeric_data),
      Min = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, min(x, na.rm = TRUE))),
      `1st_Qu.` = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, quantile(x, probs = 0.25, na.rm = TRUE))),
      Median = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, median(x, na.rm = TRUE))),
      Mean = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, mean(x, na.rm = TRUE))),
      `3rd_Qu.` = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, quantile(x, probs = 0.75, na.rm = TRUE))),
      Max = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, max(x, na.rm = TRUE))),
      SD = sapply(numeric_data, function(x) ifelse(all(is.na(x)), NA, sd(x, na.rm = TRUE)))
    )
    
    # Convert Variable column to character to prevent formatting issues
    summary_df$Variable <- as.character(summary_df$Variable)
    
    # Format numeric values for better readability
    summary_df[-1] <- lapply(summary_df[-1], function(x) round(x, 2))
    
    # Render table with better styling
    tags$div(
      class = "container-fluid",
      style = "background-color: #f8f9fa; padding: 15px; border-radius: 10px; box-shadow: 2px 2px 5px rgba(0,0,0,0.1);",
      h5("Statistical Summary"),
      DT::datatable(
        summary_df,
        options = list(
          scrollX = TRUE,
          pageLength = 10,
          dom = 't',  # Display only the table without search box, etc.
          autoWidth = FALSE,
          columnDefs = list(
            list(className = 'dt-center', targets = "_all")  # Center align all columns
          )
        ),
        rownames = FALSE
      )
    )
  })

  # Listen for the "Generate Plot" button and create the corresponding plot
  observeEvent(input$plotData, {
    req(rv$featured, input$edaActiveDataset, input$plotType)
    
    dataset_name <- input$edaActiveDataset
    data <- rv$featured[[dataset_name]]
    
    # Ensure the dataset is not empty
    req(nrow(data) > 0)
    
    # Correlation Heatmap
    if (input$plotType == "heatmap") {
      numeric_data <- data %>% dplyr::select_if(is.numeric)
      
      if (ncol(numeric_data) > 1) {
        corr_matrix <- cor(numeric_data, use = "pairwise.complete.obs")
        output$heatmapPlot <- renderPlot({
          ggcorrplot::ggcorrplot(corr_matrix, 
                                 lab = TRUE, 
                                 outline.color = "white",
                                 title = "Correlation Heatmap")
        })
      } else {
        showNotification("Heatmap requires at least two numeric features.", type = "warning")
      }
    } else {
      # Create the base ggplot and dynamically generate the title
      plot_title <- switch(input$plotType,
                           "hist" = paste("Histogram of", input$xVar),
                           "boxplot" = paste("Boxplot of", input$yVar, "by", input$xVar),
                           "bar" = paste("Bar Chart of", input$xVar),
                           "scatter" = paste("Scatter Plot of", input$xVar, "vs", input$yVar),
                           "")
      
      p <- ggplot(data, aes_string(x = input$xVar))
      
      # Histogram
      if (input$plotType == "hist") {
        p <- p + geom_histogram(alpha = input$alpha, fill = "blue", bins = 30)
        
        # Boxplot
      } else if (input$plotType == "boxplot") {
        req(input$yVar)  # Requires Y variable
        p <- p + geom_boxplot(aes_string(y = input$yVar), alpha = input$alpha)
        
        # Bar Chart
      } else if (input$plotType == "bar") {
        p <- p + geom_bar(stat = "count", fill = "blue", alpha = input$alpha)
        
        # Scatter Plot
      } else if (input$plotType == "scatter") {
        req(input$yVar)  # Requires Y variable
        p <- p + geom_point(aes_string(y = input$yVar), alpha = input$alpha, color = "blue")
      }
      
      # Add title
      p <- p + ggtitle(plot_title) +
        theme(plot.title = element_text(hjust = 0.5, face = "bold", size = 14))
      
      # Render ggplotly chart
      output$edaPlot <- renderPlotly({ ggplotly(p) })
    }
  })
  
  # Monitor `edaActiveDataset` and update the selection dropdown
  observe({
    req(rv$featured)
    updateSelectInput(session, "edaActiveDataset", choices = names(rv$featured), selected = names(rv$featured)[1])
  })
  
  
  
  # 2.5 Download & Reset Server
  # Monitor changes in `activeDataset` and update the dataset selection dropdown for downloading
  observe({
    req(rv$featured)
    updateSelectInput(session, "activeDatasetDL", choices = names(rv$featured), selected = names(rv$featured)[1])
  })
  
  # Handle data download logic
  output$downloadData <- downloadHandler(
    filename = function() {
      dataset_name <- input$activeDatasetDL
      file_ext <- input$downloadFormat
      
      paste0(dataset_name, ".", file_ext)
    },
    
    content = function(file) {
      req(rv$featured, input$activeDatasetDL, input$downloadFormat)
      
      dataset_name <- input$activeDatasetDL
      data <- rv$featured[[dataset_name]]
      
      switch(input$downloadFormat,
             "csv" = write.csv(data, file, row.names = FALSE),
             "xlsx" = writexl::write_xlsx(data, file),
             "json" = jsonlite::write_json(data, file, pretty = TRUE, auto_unbox = TRUE),
             "rds" = saveRDS(data, file)
      )
    }
  )
  
  # Monitor Reset button and trigger application reset
  observeEvent(input$reset, {
    showModal(modalDialog(
      title = "Confirm Reset",
      "Are you sure you want to reset the application? All uploaded and processed datasets will be cleared.",
      footer = tagList(
        modalButton("Cancel"),
        actionButton("confirmReset", "Yes, Reset", class = "btn-danger")
      )
    ))
  })
  
  # Handle Reset confirmation
  observeEvent(input$confirmReset, {
    removeModal()
    session$reload()
  })
}





# 3. Run the application
shinyApp(ui, server)
