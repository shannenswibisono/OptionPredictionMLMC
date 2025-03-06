# Load the necessary libraries
library(readxl)
library(dplyr)
library(lubridate)
library(xgboost)
library(caret)
library(openxlsx)

############## CALLLLL ###############
# Read training and testing data for S&P 500 Call options
train_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Training Data")
test_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Testing Data")

# Prepare features and target for training
x_train_xgb <- as.matrix(train_data_xgb[, c('Price.', 'Maturity', 'Strike', 'RVlag1', 'RVlag7', 'RVlag15', 'RVlag30', 
                                            'BVlag1', 'BVlag7', 'BVlag15', 'BVlag30', 'SJVlag7', 'SJVlag15', 'SJVlag30')])
y_train_xgb <- train_data_xgb[['Last']]

# Prepare features and target for testing
x_test_xgb <- as.matrix(test_data_xgb[, c('Price.', 'Maturity', 'Strike', 'RVlag1', 'RVlag7', 'RVlag15', 'RVlag30', 
                                          'BVlag1', 'BVlag7', 'BVlag15', 'BVlag30', 'SJVlag7', 'SJVlag15', 'SJVlag30')])
y_test_xgb <- test_data_xgb[['Last']]

# Set up parameters for the XGBoost model
params <- list(
  base_score = mean(y_train),          # Initialize with mean target value
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  eta = 0.05,                          # Equivalent to learning_rate
  objective = "reg:tweedie",           # Use Tweedie regression
  tweedie_p = 1.5,        # Specify the Tweedie p value
  max_bin = 256,                       # Adjust this based on your tree method
  max_depth = 7,                       # Depth of the tree
  min_child_weight = 1,                # Minimum sum of instance weight
  gamma = 0.01,                        # Minimum loss reduction (equivalent to min_split_loss)
  missing = NA,                        # Missing value marker (use NA for R)
  n_estimators = 2500,                 # Number of boosting rounds
  num_parallel_tree = 1,
  #random_state = 1234,                         # Equivalent to random_state
  reg_alpha = 0,                       # L1 regularization term
  reg_lambda = 1,                      # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",                # Automatically selects tree-growing algorithm
  max_delta_step = 0                   # Constraint for update step
)

# Convert training and testing sets to DMatrix format for XGBoost
dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb, missing = params$missing)
dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb, missing = params$missing)

# Train the model with the specified parameters
xgb_final_model <- xgb.train(
  params = params,
  data = dtrain,
  nrounds = params$n_estimators,
  watchlist = list(train = dtrain, test = dtest),
  early_stopping_rounds = 10
)

# Make predictions and calculate RMSE and MAPE
y_pred <- predict(xgb_final_model, dtest)
rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
cat("Test RMSE with specified XGBoost model:", rmse, "\n")

mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
cat("Test MAPE with specified XGBoost model:", mape, "\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data_xgb$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_xgb,
  Predicted = y_pred
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost DataRobot SPCall.xlsx"
write.xlsx(results, output_file)

# Make predictions and calculate RMSE and MAPE Training Data
y_pred_train <- predict(xgb_final_model, dtrain)
rmse_train <- sqrt(mean((y_train_xgb - y_pred_train)^2))
cat("Train RMSE with specified XGBoost model:", rmse_train, "\n")

mape_train <- mean(abs((y_pred_train - y_train_xgb) / y_train_xgb)) * 100
cat("Train MAPE with specified XGBoost model:", mape_train, "\n")

# Prepare the data for saving
resultstrainspcall <- data.frame(
  Date = train_data_xgb$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_train_xgb,
  Predicted = y_pred_train
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost TRAIN DataRobot SPCall.xlsx"
write.xlsx(resultstrainspcall, output_file)
cat("Results saved to Excel file:", output_file, "\n")

#### CEK GRIDSEARCH ####
# Load necessary libraries
library(xgboost)
library(caret)

# Define a small grid for testing, including the required nrounds and gamma columns
tune_grid <- expand.grid(
  nrounds = c(100, 2500),               # Small number of rounds for testing
  max_depth = c(6, 7, 8),             # Adjusted max_depth to fit provided value
  eta = c(0.05, 0.01),                 # Adjusted learning rate range to fit provided value
  gamma = c(0, 0.1),                  # Variation for gamma
  colsample_bytree = c(0.5, 0.6),     # Adjusted colsample_bytree to include provided value
  min_child_weight = c(1, 2),         # Test min_child_weight variations
  subsample = c(1)                    # Adjusted subsample to fit provided value
)

# Define train control
train_control <- trainControl(
  method = "cv",                       # Cross-validation
  number = 5,                          # 5-fold CV
  verboseIter = TRUE                   # Print progress
)

# Run grid search using caret's train function
xgb_tune <- train(
  x = x_train_xgb,                    # Ensure to use x_train_xgb as previously defined
  y = y_train_xgb,                    # Ensure to use y_train_xgb as previously defined
  method = "xgbTree",
  trControl = train_control,
  tuneGrid = tune_grid,
  metric = "RMSE"
)

# Print the best parameters found
print(xgb_tune$bestTune)

# Using best parameters from grid search to retrain model on full training data
best_params <- list(
  objective = "reg:tweedie",               # Use tweedie as specified
  eval_metric = "rmse",
  base_margin_initialize = TRUE,        # Set base_margin_initialize as specified
  colsample_bylevel = 1,                # Set colsample_bylevel as specified
  colsample_bytree = xgb_tune$bestTune$colsample_bytree,  # Use the best from grid search
  learning_rate = 0.05,                # Set learning_rate as specified
  loss = "tweedie",                    # Loss function as specified
  max_bin = 256,                       # Set max_bin as specified (note: only for certain algorithms)
  max_depth = xgb_tune$bestTune$max_depth,  # Use best max_depth found
  min_child_weight = xgb_tune$bestTune$min_child_weight,  # Use best min_child_weight found
  min_split_loss = 0.01,               # Set min_split_loss as specified
  missing_value = 0,                   # Set missing_value as specified
  n_estimators = 2500,                 # Set n_estimators as specified
  num_parallel_tree = 1,               # Set num_parallel_tree as specified
  random_state = 1234,                 # Set random_state as specified
  reg_alpha = 0,                       # Set reg_alpha as specified
  reg_lambda = 1,                      # Set reg_lambda as specified
  scale_pos_weight = 1,                # Set scale_pos_weight as specified
  smooth_interval = 200,               # Set smooth_interval as specified
  subsample = 1,                       # Set subsample as specified
  tree_method = "auto",                # Set tree_method as specified
  tweedie_p = 1.5                     # Set tweedie_p as specified
)

# Convert data to DMatrix format for XGBoost
dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb, missing = 0)
dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb, missing = 0)

# Train final model with selected parameters
xgb_final_model <- xgb.train(
  params = best_params,
  data = dtrain,
  nrounds = 2500,                     # Use the number of rounds specified
  watchlist = list(train = dtrain, test = dtest),
  early_stopping_rounds = 10
)

# Predictions and error calculations
y_pred <- predict(xgb_final_model, dtest)
rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
cat("Test RMSE with tuned model:", rmse, "\n")

mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
cat("Test MAPE with tuned model:", mape, "\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data_xgb$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_xgb,
  Predicted = y_pred
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost GridSCV SPCall.xlsx"
write.xlsx(results, output_file)

cat("Results saved to Excel file:", output_file, "\n")



##### Call dengan masing2 lag ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Training Data")
test_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Testing Data")

# Set up parameters for the XGBoost model
params <- list(
  base_score = mean(y_train),          # Initialize with mean target value
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  eta = 0.05,                          # Equivalent to learning_rate
  objective = "reg:tweedie",           # Use Tweedie regression
  tweedie_p = 1.5,        # Specify the Tweedie p value
  max_bin = 256,                       # Adjust this based on your tree method
  max_depth = 7,                       # Depth of the tree
  min_child_weight = 1,                # Minimum sum of instance weight
  gamma = 0.01,                        # Minimum loss reduction (equivalent to min_split_loss)
  missing = NA,                        # Missing value marker (use NA for R)
  n_estimators = 2500,                 # Number of boosting rounds
  num_parallel_tree = 1,
  #random_state = 1234,                         # Equivalent to random_state
  reg_alpha = 0,                       # L1 regularization term
  reg_lambda = 1,                      # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",                # Automatically selects tree-growing algorithm
  max_delta_step = 0                   # Constraint for update step
)

# Define the lag configurations
lags <- list(
  lag1 = c("RVlag1", "BVlag1"),
  lag7 = c("RVlag7", "BVlag7", "SJVlag7"),
  lag15 = c("RVlag15", "BVlag15", "SJVlag15"),
  lag30 = c("RVlag30", "BVlag30", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_SPCall_Lags.xlsx"
wb <- createWorkbook()

# Loop through each lag configuration
for (lag_name in names(lags)) {
  cat("\nTesting with lag:", lag_name, "\n")
  
  # Get the relevant lag features
  selected_features <- c("Price.", "Maturity", "Strike", lags[[lag_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data_xgb %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data_xgb[['Last']]
  
  x_test_xgb <- test_data_xgb %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data_xgb[['Last']]
  
  # Convert data to DMatrix format
  dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb)
  dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb)
  
  # Train the model
  xgb_model <- xgb.train(
    params = params,
    data = dtrain,
    nrounds = params$n_estimators,
    watchlist = list(train = dtrain, test = dtest),
    #early_stopping_rounds = 10,
  )
  
  # Make predictions and calculate RMSE and MAPE
  y_pred <- predict(xgb_model, dtest)
  rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
  mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
  
  cat("Lag:", lag_name, "RMSE:", rmse, "MAPE:", mape, "\n")
  
  # Prepare results for saving
  results <- data.frame(
    Date = test_data_xgb$Time,  # Assuming 'Time' column exists in the test data
    Actual = y_test_xgb,
    Predicted = y_pred,
    RMSE = rmse,
    MAPE = mape
  )
  
  # Add results to a new sheet in the workbook
  addWorksheet(wb, sheetName = lag_name)
  writeData(wb, sheet = lag_name, x = results)
}

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("All results saved to Excel file:", output_file, "\n")


##### Call masing2 dengan variation yang berbeda ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Training Data")
test_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Testing Data")

# Set up parameters for the XGBoost model
params <- list(
  base_score = mean(y_train),          # Initialize with mean target value
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  eta = 0.05,                          # Equivalent to learning_rate
  objective = "reg:tweedie",           # Use Tweedie regression
  tweedie_p = 1.5,        # Specify the Tweedie p value
  max_bin = 256,                       # Adjust this based on your tree method
  max_depth = 7,                       # Depth of the tree
  min_child_weight = 1,                # Minimum sum of instance weight
  gamma = 0.01,                        # Minimum loss reduction (equivalent to min_split_loss)
  missing = NA,                        # Missing value marker (use NA for R)
  n_estimators = 2500,                 # Number of boosting rounds
  num_parallel_tree = 1,
  #random_state = 1234,                         # Equivalent to random_state
  reg_alpha = 0,                       # L1 regularization term
  reg_lambda = 1,                      # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",                # Automatically selects tree-growing algorithm
  max_delta_step = 0                   # Constraint for update step
)

# Define the feature configurations based on variations
variations <- list(
  RV = c("RVlag1", "RVlag7", "RVlag15", "RVlag30"),
  BV = c("BVlag1", "BVlag7", "BVlag15", "BVlag30"),
  SJV = c("SJVlag1","SJVlag7", "SJVlag15", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_SPCall_Variations.xlsx"
wb <- createWorkbook()

# Loop through each variation
for (var_name in names(variations)) {
  cat("\nTesting with variation:", var_name, "\n")
  
  # Get the relevant features for the current variation
  selected_features <- c("Price.", "Maturity", "Strike", variations[[var_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data_xgb %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data_xgb[['Last']]
  
  x_test_xgb <- test_data_xgb %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data_xgb[['Last']]
  
  # Convert data to DMatrix format
  dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb)
  dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb)
  
  # Train the model
  xgb_model <- xgb.train(
    params = params,
    data = dtrain,
    nrounds = params$n_estimators,
    watchlist = list(train = dtrain, test = dtest)
  )
  
  # Make predictions and calculate RMSE and MAPE
  y_pred <- predict(xgb_model, dtest)
  rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
  mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
  
  cat("Variation:", var_name, "RMSE:", rmse, "MAPE:", mape, "\n")
  
  # Prepare results for saving
  results <- data.frame(
    Date = test_data_xgb$Time,  # Assuming 'Time' column exists in the test data
    Actual = y_test_xgb,
    Predicted = y_pred,
    RMSE = rmse,
    MAPE = mape
  )
  
  # Add results to a new sheet in the workbook
  addWorksheet(wb, sheetName = var_name)
  writeData(wb, sheet = var_name, x = results)
}

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("All results saved to Excel file:", output_file, "\n")




##### Call dengan semua data ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Training Data")
test_data_xgb <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx", sheet = "Testing Data")

# Set up parameters for the XGBoost model
params <- list(
  base_score = mean(y_train),          # Initialize with mean target value
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  eta = 0.05,                          # Equivalent to learning_rate
  objective = "reg:tweedie",           # Use Tweedie regression
  tweedie_p = 1.5,        # Specify the Tweedie p value
  max_bin = 256,                       # Adjust this based on your tree method
  max_depth = 7,                       # Depth of the tree
  min_child_weight = 1,                # Minimum sum of instance weight
  gamma = 0.01,                        # Minimum loss reduction (equivalent to min_split_loss)
  missing = NA,                        # Missing value marker (use NA for R)
  n_estimators = 2500,                 # Number of boosting rounds
  num_parallel_tree = 1,
  #random_state = 1234,                         # Equivalent to random_state
  reg_alpha = 0,                       # L1 regularization term
  reg_lambda = 1,                      # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",                # Automatically selects tree-growing algorithm
  max_delta_step = 0                   # Constraint for update step
)

# Define the lag configurations
lags <- list(
  RVlag1 = c("RVlag1"),
  RVlag7 = c("RVlag7"),
  RVlag15 = c("RVlag15"),
  RVlag30 = c("RVlag30"),
  BVlag1 = c("BVlag1"),
  BVlag7 = c("BVlag7"),
  BVlag15 = c("BVlag15"),
  BVlag30 = c("BVlag30"),
  SJVlag7 = c("SJVlag7"),
  SJVlag15 = c("SJVlag15"),
  SJVlag30 = c("SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_SPCall_All.xlsx"
wb <- createWorkbook()

# Initialize summary dataframe
summary_results <- data.frame(
  Lag = character(),
  RMSE = numeric(),
  MAPE = numeric(),
  stringsAsFactors = FALSE
)

# Loop through each lag configuration
for (lag_name in names(lags)) {
  cat("\nTesting with lag:", lag_name, "\n")
  
  # Get the relevant lag features
  selected_features <- c("Price.", "Maturity", "Strike", lags[[lag_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data_xgb %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data_xgb[['Last']]
  
  x_test_xgb <- test_data_xgb %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data_xgb[['Last']]
  
  # Convert data to DMatrix format
  dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb)
  dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb)
  
  # Train the model
  xgb_model <- xgb.train(
    params = params,
    data = dtrain,
    nrounds = params$n_estimators,
    watchlist = list(train = dtrain, test = dtest),
    #early_stopping_rounds = 10,
  )
  
  # Make predictions and calculate RMSE and MAPE
  y_pred <- predict(xgb_model, dtest)
  rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
  mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
  
  cat("Lag:", lag_name, "RMSE:", rmse, "MAPE:", mape, "\n")
  
  # Append to summary dataframe
  summary_results <- rbind(summary_results, data.frame(
    Lag = lag_name,
    RMSE = rmse,
    MAPE = mape
  ))
  
  # Prepare results for saving
  results <- data.frame(
    Date = test_data_xgb$Time,  # Assuming 'Time' column exists in the test data
    Actual = y_test_xgb,
    Predicted = y_pred,
    RMSE = rmse,
    MAPE = mape
  )
  
  # Add results to a new sheet in the workbook
  addWorksheet(wb, sheetName = lag_name)
  writeData(wb, sheet = lag_name, x = results)
}

# Add summary results to the workbook
addWorksheet(wb, sheetName = "Summary")
writeData(wb, sheet = "Summary", x = summary_results)

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("All results saved to Excel file:", output_file, "\n")

############## PUT ###############
# Load necessary libraries
library(xgboost)
library(caret)
library(readxl)

# Load Put option training and testing data
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Testing Data Put")

# Prepare features and target for training and testing (Put)
x_train_lgb_put <- as.matrix(train_data_put[, c('Price.', 'Maturity', 'Strike', 'RVlag1', 'RVlag7', 'RVlag15', 'RVlag30', 
                                                'BVlag1', 'BVlag7', 'BVlag15', 'BVlag30', 'SJVlag7', 'SJVlag15', 'SJVlag30')])
y_train_lgb_put <- train_data_put[['Last']]
x_test_lgb_put <- as.matrix(test_data_put[, c('Price.', 'Maturity', 'Strike', 'RVlag1', 'RVlag7', 'RVlag15', 'RVlag30', 
                                              'BVlag1', 'BVlag7', 'BVlag15', 'BVlag30', 'SJVlag7', 'SJVlag15', 'SJVlag30')])
y_test_lgb_put <- test_data_put[['Last']]

# Set up parameters for the XGBoost model
params <- list(
  objective = "count:poisson",            # Using Poisson distribution for Put options
  eval_metric = "rmse",
  base_score = mean(y_train_lgb_put),     # Equivalent to initializing base margin
  booster = "gbtree",
  colsample_bylevel = 1,
  colsample_bytree = 1,
  eta = 0.02,                             # Learning rate
  max_bin = 256,
  max_delta_step = 0,
  max_depth = 10,
  min_child_weight = 1,
  gamma = 0.01,                           # Min split loss
  n_estimators = 6250,
  num_parallel_tree = 1,
  seed = 1234,
  reg_alpha = 0,
  reg_lambda = 1,
  subsample = 1,
  tree_method = "auto"
)

# Convert training and testing sets to DMatrix format
dtrain <- xgb.DMatrix(data = x_train_lgb_put, label = y_train_lgb_put, missing = 0)
dtest <- xgb.DMatrix(data = x_test_lgb_put, label = y_test_lgb_put, missing = 0)

# Train the model with specified parameters
xgb_final_model <- xgb.train(
  params = params,
  data = dtrain,
  nrounds = params$n_estimators,
  watchlist = list(train = dtrain, test = dtest),
  early_stopping_rounds = 10
)

# Make predictions and calculate RMSE and MAPE
y_pred <- predict(xgb_final_model, dtest)
rmse <- sqrt(mean((y_test_lgb_put - y_pred)^2))
cat("Test RMSE with specified model:", rmse, "\n")

mape <- mean(abs((y_pred - y_test_lgb_put) / y_test_lgb_put)) * 100
cat("Test MAPE with specified model:", mape, "\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data_put$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_lgb_put,
  Predicted = y_pred
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost DataRobot SPPut.xlsx"
write.xlsx(results, output_file)
cat("Results saved to Excel file:", output_file, "\n")

# Make predictions and calculate RMSE and MAPE on Training
y_pred_puttrain <- predict(xgb_final_model, dtrain)
rmse_puttrain <- sqrt(mean((y_train_lgb_put - y_pred_puttrain)^2))
cat("Train RMSE with specified model:", rmse_puttrain, "\n")

mape_puttrain <- mean(abs((y_pred_puttrain - y_train_lgb_put) / y_train_lgb_put)) * 100
cat("Train MAPE with specified model:", mape_puttrain, "\n")

# Prepare the data for saving
resultsputtrain <- data.frame(
  Date = train_data_put$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_train_lgb_put,
  Predicted = y_pred_puttrain
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost TRAIN DataRobot SPPut.xlsx"
write.xlsx(resultsputtrain, output_file)
cat("Results saved to Excel file:", output_file, "\n")

#### Dengan GridSearch ####
# Grid Search Setup for Parameter Tuning
# Define a small grid for testing
tune_grid <- expand.grid(
  nrounds = c(50, 100),                  # Limited rounds for testing 20,100
  max_depth = c(8, 10),
  eta = c(0.01, 0.02),
  gamma = c(0, 0.01),
  colsample_bytree = c(0.8, 1),
  min_child_weight = c(1, 2),
  subsample = c(0.8, 1)
)

# Define train control for cross-validation
train_control <- trainControl(
  method = "cv",
  number = 5,
  verboseIter = TRUE
)

# Run grid search using caret's train function
xgb_tune <- train(
  x = x_train_lgb_put, 
  y = y_train_lgb_put,
  method = "xgbTree",
  trControl = train_control,
  tuneGrid = tune_grid,
  metric = "RMSE"
)

# Print best parameters from grid search
print(xgb_tune$bestTune)

# Using best parameters from grid search to retrain model
best_params <- list(
  objective = "count:poisson",
  eval_metric = "rmse",
  max_depth = xgb_tune$bestTune$max_depth,
  eta = xgb_tune$bestTune$eta,
  colsample_bytree = xgb_tune$bestTune$colsample_bytree,
  subsample = xgb_tune$bestTune$subsample,
  min_child_weight = xgb_tune$bestTune$min_child_weight,
  tree_method = "auto",
  n_estimators = 6250,
  base_score = mean(y_train_lgb_put)
)

# Retrain final model with best parameters
xgb_final_model <- xgb.train(
  params = best_params,
  data = dtrain,
  nrounds = 6250,
  watchlist = list(train = dtrain, test = dtest),
  early_stopping_rounds = 10
)

# Make predictions and calculate RMSE and MAPE with tuned model
y_pred <- predict(xgb_final_model, dtest)
rmse <- sqrt(mean((y_test_lgb_put - y_pred)^2))
cat("Test RMSE with tuned model:", rmse, "\n")

mape <- mean(abs((y_pred - y_test_lgb_put) / y_test_lgb_put)) * 100
cat("Test MAPE with tuned model:", mape, "\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data_put$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_lgb_put,
  Predicted = y_pred
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost GridSCV SPPut.xlsx"
write.xlsx(results, output_file)

cat("Results saved to Excel file:", output_file, "\n")


##### Put dengan beda2 lag ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Testing Data Put")

# Set up parameters for the XGBoost model
params <- list(
  objective = "count:poisson",            # Using Poisson distribution for Put options
  eval_metric = "rmse",
  base_score = mean(y_train_lgb_put),     # Equivalent to initializing base margin
  booster = "gbtree",
  colsample_bylevel = 1,
  colsample_bytree = 1,
  eta = 0.02,                             # Learning rate
  max_bin = 256,
  max_delta_step = 0,
  max_depth = 10,
  min_child_weight = 1,
  gamma = 0.01,                           # Min split loss
  n_estimators = 6250,
  num_parallel_tree = 1,
  seed = 1234,
  reg_alpha = 0,
  reg_lambda = 1,
  subsample = 1,
  tree_method = "auto"
)

# Define the lag configurations
lags <- list(
  lag1 = c("RVlag1", "BVlag1"),
  lag7 = c("RVlag7", "BVlag7", "SJVlag7"),
  lag15 = c("RVlag15", "BVlag15", "SJVlag15"),
  lag30 = c("RVlag30", "BVlag30", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_SPPut_Lags.xlsx"
wb <- createWorkbook()

# Loop through each lag configuration
for (lag_name in names(lags)) {
  cat("\nTesting with lag:", lag_name, "\n")
  
  # Get the relevant lag features
  selected_features <- c("Price.", "Maturity", "Strike", lags[[lag_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data_put %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data_put[['Last']]
  
  x_test_xgb <- test_data_put %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data_put[['Last']]
  
  # Convert data to DMatrix format
  dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb)
  dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb)
  
  # Train the model
  xgb_model <- xgb.train(
    params = params,
    data = dtrain,
    nrounds = params$n_estimators,
    watchlist = list(train = dtrain, test = dtest),
    #early_stopping_rounds = 10,
  )
  
  # Make predictions and calculate RMSE and MAPE
  y_pred <- predict(xgb_model, dtest)
  rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
  mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
  
  cat("Lag:", lag_name, "RMSE:", rmse, "MAPE:", mape, "\n")
  
  # Prepare results for saving
  results <- data.frame(
    Date = test_data_put$Time,  # Assuming 'Time' column exists in the test data
    Actual = y_test_xgb,
    Predicted = y_pred,
    RMSE = rmse,
    MAPE = mape
  )
  
  # Add results to a new sheet in the workbook
  addWorksheet(wb, sheetName = lag_name)
  writeData(wb, sheet = lag_name, x = results)
}

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("All results saved to Excel file:", output_file, "\n")


##### Put masing2 dengan variation yang berbeda ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Testing Data Put")

# Set up parameters for the XGBoost model
params <- list(
  objective = "count:poisson",            # Using Poisson distribution for Put options
  eval_metric = "rmse",
  base_score = mean(y_train_lgb_put),     # Equivalent to initializing base margin
  booster = "gbtree",
  colsample_bylevel = 1,
  colsample_bytree = 1,
  eta = 0.02,                             # Learning rate
  max_bin = 256,
  max_delta_step = 0,
  max_depth = 10,
  min_child_weight = 1,
  gamma = 0.01,                           # Min split loss
  n_estimators = 6250,
  num_parallel_tree = 1,
  seed = 1234,
  reg_alpha = 0,
  reg_lambda = 1,
  subsample = 1,
  tree_method = "auto"
)

# Define the feature configurations based on variations
variations <- list(
  RV = c("RVlag1", "RVlag7", "RVlag15", "RVlag30"),
  BV = c("BVlag1", "BVlag7", "BVlag15", "BVlag30"),
  SJV = c("SJVlag1","SJVlag7", "SJVlag15", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_SPPut_Variations.xlsx"
wb <- createWorkbook()

# Loop through each variation
for (var_name in names(variations)) {
  cat("\nTesting with variation:", var_name, "\n")
  
  # Get the relevant features for the current variation
  selected_features <- c("Price.", "Maturity", "Strike", variations[[var_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data_put %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data_put[['Last']]
  
  x_test_xgb <- test_data_put %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data_put[['Last']]
  
  # Convert data to DMatrix format
  dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb)
  dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb)
  
  # Train the model
  xgb_model <- xgb.train(
    params = params,
    data = dtrain,
    nrounds = params$n_estimators,
    watchlist = list(train = dtrain, test = dtest)
  )
  
  # Make predictions and calculate RMSE and MAPE
  y_pred <- predict(xgb_model, dtest)
  rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
  mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
  
  cat("Variation:", var_name, "RMSE:", rmse, "MAPE:", mape, "\n")
  
  # Prepare results for saving
  results <- data.frame(
    Date = test_data_put$Time,  # Assuming 'Time' column exists in the test data
    Actual = y_test_xgb,
    Predicted = y_pred,
    RMSE = rmse,
    MAPE = mape
  )
  
  # Add results to a new sheet in the workbook
  addWorksheet(wb, sheetName = var_name)
  writeData(wb, sheet = var_name, x = results)
}

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("All results saved to Excel file:", output_file, "\n")


##### Put dengan beda2 lag ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx", sheet = "Testing Data Put")

# Set up parameters for the XGBoost model
params <- list(
  objective = "count:poisson",            # Using Poisson distribution for Put options
  eval_metric = "rmse",
  base_score = mean(y_train_lgb_put),     # Equivalent to initializing base margin
  booster = "gbtree",
  colsample_bylevel = 1,
  colsample_bytree = 1,
  eta = 0.02,                             # Learning rate
  max_bin = 256,
  max_delta_step = 0,
  max_depth = 10,
  min_child_weight = 1,
  gamma = 0.01,                           # Min split loss
  n_estimators = 6250,
  num_parallel_tree = 1,
  seed = 1234,
  reg_alpha = 0,
  reg_lambda = 1,
  subsample = 1,
  tree_method = "auto"
)

# Define the lag configurations
lags <- list(
  RVlag1 = c("RVlag1"),
  RVlag7 = c("RVlag7"),
  RVlag15 = c("RVlag15"),
  RVlag30 = c("RVlag30"),
  BVlag1 = c("BVlag1"),
  BVlag7 = c("BVlag7"),
  BVlag15 = c("BVlag15"),
  BVlag30 = c("BVlag30"),
  SJVlag7 = c("SJVlag7"),
  SJVlag15 = c("SJVlag15"),
  SJVlag30 = c("SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_SPPut_All.xlsx"
wb <- createWorkbook()

# Initialize summary dataframe
summary_results <- data.frame(
  Lag = character(),
  RMSE = numeric(),
  MAPE = numeric(),
  stringsAsFactors = FALSE
)

# Loop through each lag configuration
for (lag_name in names(lags)) {
  cat("\nTesting with lag:", lag_name, "\n")
  
  # Get the relevant lag features
  selected_features <- c("Price.", "Maturity", "Strike", lags[[lag_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data_put %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data_put[['Last']]
  
  x_test_xgb <- test_data_put %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data_put[['Last']]
  
  # Convert data to DMatrix format
  dtrain <- xgb.DMatrix(data = x_train_xgb, label = y_train_xgb)
  dtest <- xgb.DMatrix(data = x_test_xgb, label = y_test_xgb)
  
  # Train the model
  xgb_model <- xgb.train(
    params = params,
    data = dtrain,
    nrounds = params$n_estimators,
    watchlist = list(train = dtrain, test = dtest),
    #early_stopping_rounds = 10,
  )
  
  # Make predictions and calculate RMSE and MAPE
  y_pred <- predict(xgb_model, dtest)
  rmse <- sqrt(mean((y_test_xgb - y_pred)^2))
  mape <- mean(abs((y_pred - y_test_xgb) / y_test_xgb)) * 100
  
  cat("Lag:", lag_name, "RMSE:", rmse, "MAPE:", mape, "\n")
  
  # Append to summary dataframe
  summary_results <- rbind(summary_results, data.frame(
    Lag = lag_name,
    RMSE = rmse,
    MAPE = mape
  ))
  
  # Prepare results for saving
  results <- data.frame(
    Date = test_data_put$Time,  # Assuming 'Time' column exists in the test data
    Actual = y_test_xgb,
    Predicted = y_pred,
    RMSE = rmse,
    MAPE = mape
  )
  
  # Add results to a new sheet in the workbook
  addWorksheet(wb, sheetName = lag_name)
  writeData(wb, sheet = lag_name, x = results)
}

# Add summary results to the workbook
addWorksheet(wb, sheetName = "Summary")
writeData(wb, sheet = "Summary", x = summary_results)

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)
cat("All results saved to Excel file:", output_file, "\n")
