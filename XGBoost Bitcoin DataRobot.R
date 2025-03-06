# Load the necessary libraries
library(readxl)
library(dplyr)
library(lubridate)
library(xgboost)
library(caret)
library(openxlsx)

##### CALLLL #######
# Load the training and testing data
train_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Training Data Call")
test_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Testing Data Call")

# Prepare features and target for training
library(dplyr)
x_train_xgb <- train_data %>%
  select('Asset Price', Maturity, Strike, RVlag1, RVlag7, RVlag15, RVlag30, 
         BVlag1, BVlag7, BVlag15, BVlag30, SJVlag7, SJVlag15, SJVlag30) %>%
  as.matrix()

y_train_xgb <- train_data[['Last']]

# Prepare features and target for testing
x_test_xgb <- test_data %>%
  select('Asset Price', Maturity, Strike, RVlag1, RVlag7, RVlag15, RVlag30, 
         BVlag1, BVlag7, BVlag15, BVlag30, SJVlag7, SJVlag15, SJVlag30) %>%
  as.matrix()
y_test_xgb <- test_data[['Last']]

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  learning_rate = 0.05,             # This is the same as 'eta'
  objective = "reg:tweedie",        # Use Tweedie regression
  tweedie_p = 1.5,                  # Specify the Tweedie p value
  max_bin = 256,                    # Adjust this based on your tree method
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 2500,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,              # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",              # Set to automatic
  max_delta_step = 0,                # Constraint for update step
  interval = 10,                     # Watchlist logging interval (not a standard parameter)
  smooth_interval = 200               # Smooth interval (not a standard parameter)
)

# Set the seed for reproducibility
set.seed(params$random_state)

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

# Make predictions and calculate RMSE and MAPE on training data
y_pred_train <- predict(xgb_final_model, dtrain)
rmse_train <- sqrt(mean((y_train_xgb - y_pred_train)^2))
cat("Train RMSE with specified XGBoost model:", rmse_train, "\n")

mape_train <- mean(abs((y_pred_train - y_train_xgb) / y_train_xgb)) * 100
cat("Train MAPE with specified XGBoost model:", mape_train, "\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_xgb,
  Predicted = y_pred
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost DataRobot BTCCall.xlsx"
write.xlsx(results, output_file)

cat("Results saved to Excel file:", output_file, "\n")

# Training data
results_train <- data.frame(
  Date = train_data$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_train_xgb,
  Predicted = y_pred_train
)

output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost TRAIN DataRobot BTCCall.xlsx"
write.xlsx(results_train, output_file)

# Load required libraries
library(xgboost)
library(caret)

#### GRIDSEARCHCV ####
# Define a small grid for testing (optimized for speed)
tune_grid <- expand.grid(
  nrounds = c(2500),                # Number of boosting rounds
  max_depth = c(3, 5),                   # Depth of the tree
  eta = c(0.05, 0.1),                    # Learning rate
  gamma = c(0.01, 0.1),                  # Minimum loss reduction
  colsample_bytree = c(0.5, 0.7),        # Fraction of features used per tree
  min_child_weight = c(1, 3),            # Minimum sum of instance weight
  subsample = c(1)                       # Fraction of samples used per tree
)

# Define train control for GridSearchCV
train_control <- trainControl(
  method = "cv",                         # Cross-validation
  number = 5,                            # 3-fold CV for faster processing
  verboseIter = TRUE                     # Print progress
)

# Run grid search using caret's train function
xgb_tune <- train(
  x = x_train_xgb,                       # Features for training
  y = y_train_xgb,                       # Target for training
  method = "xgbTree",                    # XGBoost method
  trControl = train_control,             # Train control settings
  tuneGrid = tune_grid,                  # Grid of parameters to search
  metric = "RMSE"                        # Optimize for RMSE
)

# Print the best parameters found by GridSearchCV
cat("Best Parameters from GridSearchCV:\n")
print(xgb_tune$bestTune)

# Extract the best parameters from GridSearchCV
best_params <- list(
  max_depth = xgb_tune$bestTune$max_depth,
  eta = xgb_tune$bestTune$eta,
  gamma = xgb_tune$bestTune$gamma,
  colsample_bytree = xgb_tune$bestTune$colsample_bytree,
  min_child_weight = xgb_tune$bestTune$min_child_weight,
  subsample = xgb_tune$bestTune$subsample,
  objective = "reg:tweedie",             # Use Tweedie regression
  tweedie_variance_power = 1.5,          # Specify the Tweedie p value
  eval_metric = "rmse",
  base_score = mean(y_train_xgb),        # Initialize with mean target value
  #seed = 1234                            # For reproducibility
)

# Train the model with the best parameters from GridSearchCV
xgb_model_grid <- xgb.train(
  params = best_params,
  data = dtrain,
  nrounds = 2500,   # Use the best number of boosting rounds
  watchlist = list(train = dtrain, test = dtest),
  early_stopping_rounds = 10,            # Stop training if no improvement
  verbose = TRUE
)

# Make predictions and calculate RMSE and MAPE for GridSearchCV model
y_pred_grid <- predict(xgb_model_grid, dtest)

rmse_grid <- sqrt(mean((y_test_xgb - y_pred_grid)^2))
cat("RMSE with GridSearchCV parameters:", rmse_grid, "\n")

mape_grid <- mean(abs((y_test_xgb - y_pred_grid) / y_test_xgb)) * 100
cat("MAPE with GridSearchCV parameters:", mape_grid, "%\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_xgb,
  Predicted = y_pred_grid
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost GridSCV BTCCall.xlsx"
write.xlsx(results, output_file)

cat("Results saved to Excel file:", output_file, "\n")


#### Call masing2 dengan lag ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Training Data Call")
test_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Testing Data Call")

# Set up parameters for the XGBoost model
params <- list(
    base_margin_initialize = TRUE,
    colsample_bylevel = 1,
    colsample_bytree = 0.5,
    learning_rate = 0.05,             # This is the same as 'eta'
    objective = "reg:tweedie",        # Use Tweedie regression
    tweedie_p = 1.5,                  # Specify the Tweedie p value
    max_bin = 256,                    # Adjust this based on your tree method
    max_depth = 5,                    # Depth of the tree
    min_child_weight = 1,             # Minimum sum of instance weight
    min_split_loss = 0.01,            # Minimum loss reduction
    missing = -999,                   # Missing value marker
    n_estimators = 2500,              # Number of boosting rounds
    num_parallel_tree = 1,
    random_state = 1234,              # For reproducibility
    reg_alpha = 0,                     # L1 regularization term
    reg_lambda = 1,                    # L2 regularization term
    scale_pos_weight = 1,
    subsample = 1,
    tree_method = "auto",              # Set to automatic
    max_delta_step = 0,                # Constraint for update step
    interval = 10,                     # Watchlist logging interval (not a standard parameter)
    smooth_interval = 200               # Smooth interval (not a standard parameter)
)

# Define the lag configurations
lags <- list(
  lag1 = c("RVlag1", "BVlag1"),
  lag7 = c("RVlag7", "BVlag7", "SJVlag7"),
  lag15 = c("RVlag15", "BVlag15", "SJVlag15"),
  lag30 = c("RVlag30", "BVlag30", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_BTCCall_Lags.xlsx"
wb <- createWorkbook()

# Loop through each lag configuration
for (lag_name in names(lags)) {
  cat("\nTesting with lag:", lag_name, "\n")
  
  # Get the relevant lag features
  selected_features <- c("Asset Price", "Maturity", "Strike", lags[[lag_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data[['Last']]
  
  x_test_xgb <- test_data %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data[['Last']]
  
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
    Date = test_data$Time,  # Assuming 'Time' column exists in the test data
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
train_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Training Data Call")
test_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Testing Data Call")

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  learning_rate = 0.05,             # This is the same as 'eta'
  objective = "reg:tweedie",        # Use Tweedie regression
  tweedie_p = 1.5,                  # Specify the Tweedie p value
  max_bin = 256,                    # Adjust this based on your tree method
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 2500,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,              # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",              # Set to automatic
  max_delta_step = 0,                # Constraint for update step
  interval = 10,                     # Watchlist logging interval (not a standard parameter)
  smooth_interval = 200               # Smooth interval (not a standard parameter)
)

# Define the feature configurations based on variations
variations <- list(
  RV = c("RVlag1", "RVlag7", "RVlag15", "RVlag30"),
  BV = c("BVlag1", "BVlag7", "BVlag15", "BVlag30"),
  SJV = c("SJVlag1","SJVlag7", "SJVlag15", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_BTCCall_Variations.xlsx"
wb <- createWorkbook()

# Loop through each variation
for (var_name in names(variations)) {
  cat("\nTesting with variation:", var_name, "\n")
  
  # Get the relevant features for the current variation
  selected_features <- c("Asset Price", "Maturity", "Strike", variations[[var_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data[['Last']]
  
  x_test_xgb <- test_data %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data[['Last']]
  
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
    Date = test_data$Time,  # Assuming 'Time' column exists in the test data
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

##### Call dengan data all ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Training Data Call")
test_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx", sheet = "Testing Data Call")

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  learning_rate = 0.05,             # This is the same as 'eta'
  objective = "reg:tweedie",        # Use Tweedie regression
  tweedie_p = 1.5,                  # Specify the Tweedie p value
  max_bin = 256,                    # Adjust this based on your tree method
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 2500,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,              # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term
  scale_pos_weight = 1,
  subsample = 1,
  tree_method = "auto",              # Set to automatic
  max_delta_step = 0,                # Constraint for update step
  interval = 10,                     # Watchlist logging interval (not a standard parameter)
  smooth_interval = 200               # Smooth interval (not a standard parameter)
)

# Define lag configurations (individual testing)
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
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_BTCCall_All.xlsx"
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
  selected_features <- c("Asset Price", "Maturity", "Strike", lags[[lag_name]])
  
  # Prepare features and target for training and testing
  x_train_xgb <- train_data %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_train_xgb <- train_data[['Last']]
  
  x_test_xgb <- test_data %>%
    select(all_of(selected_features)) %>%
    as.matrix()
  
  y_test_xgb <- test_data[['Last']]
  
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
    Date = test_data$Time,  # Assuming 'Time' column exists in the test data
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


##### PUT #####
# Load required libraries
library(readxl)
library(dplyr)
library(xgboost)

# Load the training and testing data for Put options
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Testing Data Put")

# Prepare features and target for training
x_train_xgb_put <- train_data_put %>%
  select('Asset Price', Maturity, Strike, RVlag1, RVlag7, RVlag15, RVlag30, 
         BVlag1, BVlag7, BVlag15, BVlag30, SJVlag7, SJVlag15, SJVlag30) %>%
  as.matrix()

y_train_xgb_put <- train_data_put[['Last']]

# Prepare features and target for testing
x_test_xgb_put <- test_data_put %>%
  select('Asset Price', Maturity, Strike, RVlag1, RVlag7, RVlag15, RVlag30, 
         BVlag1, BVlag7, BVlag15, BVlag30, SJVlag7, SJVlag15, SJVlag30) %>%
  as.matrix()

y_test_xgb_put <- test_data_put[['Last']]

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  interval = 10,                    # Watchlist logging interval
  learning_rate = 0.02,             # This is the same as 'eta'
  objective = "reg:gamma",          # Use gamma loss
  max_bin = 256,                    # Adjust this based on your tree method
  max_delta_step = 0,               # Constraint for update step
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 6250,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,               # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term (note the corrected spelling)
  scale_pos_weight = 1,
  smooth_interval = 200,             # Smooth interval (not a standard parameter)
  subsample = 0.7,
  tree_method = "auto",              # Set to automatic
  tweedie_p = 1.5                   # Specify the Tweedie p value
)

# Set the seed for reproducibility
set.seed(params$random_state)

# Convert training and testing sets to DMatrix format for XGBoost
dtrain_put <- xgb.DMatrix(data = x_train_xgb_put, label = y_train_xgb_put, missing = params$missing)
dtest_put <- xgb.DMatrix(data = x_test_xgb_put, label = y_test_xgb_put, missing = params$missing)

# Train the model with the specified parameters
xgb_final_model_put <- xgb.train(
  params = params,
  data = dtrain_put,
  nrounds = params$n_estimators,
  watchlist = list(train = dtrain_put, test = dtest_put),
  early_stopping_rounds = 10
)

# Make predictions and calculate RMSE and MAPE
y_pred_put <- predict(xgb_final_model_put, dtest_put)
rmse_put <- sqrt(mean((y_test_xgb_put - y_pred_put)^2))
cat("Test RMSE for Put options with specified XGBoost model:", rmse_put, "\n")

mape_put <- mean(abs((y_pred_put - y_test_xgb_put) / y_test_xgb_put)) * 100
cat("Test MAPE for Put options with specified XGBoost model:", mape_put, "\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data_put$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_xgb_put,
  Predicted = y_pred_put
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost DataRobot BTCPut.xlsx"
write.xlsx(results, output_file)

cat("Results saved to Excel file:", output_file, "\n")

# Make predictions and calculate RMSE and MAPE
y_pred_put_train <- predict(xgb_final_model_put, dtrain_put)
rmse_put_train <- sqrt(mean((y_train_xgb_put - y_pred_put_train)^2))
cat("Train RMSE for Put options with specified XGBoost model:", rmse_put_train, "\n")

mape_put_train <- mean(abs((y_pred_put_train - y_train_xgb_put) / y_train_xgb_put)) * 100
cat("Train MAPE for Put options with specified XGBoost model:", mape_put_train, "\n")

# Prepare the data for saving
resultsputtrain <- data.frame(
  Date = train_data_put$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_train_xgb_put,
  Predicted = y_pred_put_train
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost TRAIN DataRobot BTCPut.xlsx"
write.xlsx(resultsputtrain, output_file)

cat("Results saved to Excel file:", output_file, "\n")

### GRIDSEARCHCV ####
# Define a small grid for GridSearchCV
tune_grid_put <- expand.grid(
  nrounds = c(6250),                # Number of boosting rounds
  max_depth = c(3, 5),                   # Depth of the tree
  eta = c(0.02, 0.05),                   # Learning rate
  gamma = c(0.01, 0.1),                 # Minimum loss reduction
  colsample_bytree = c(0.5, 0.7),        # Fraction of features used per tree
  min_child_weight = c(1, 3),            # Minimum sum of instance weight
  subsample = c(0.7)                     # Fraction of samples used per tree
)

# Define train control
train_control_put <- trainControl(
  method = "cv",                         # Cross-validation
  number = 5,                            # 5-fold CV for faster processing
  verboseIter = TRUE                     # Print progress
)

# Run grid search using caret's train function
xgb_tune_put <- train(
  x = x_train_xgb_put,                   # Features for training
  y = y_train_xgb_put,                   # Target for training
  method = "xgbTree",                    # XGBoost method
  trControl = train_control_put,         # Train control settings
  tuneGrid = tune_grid_put,              # Grid of parameters to search
  metric = "RMSE"                        # Optimize for RMSE
)

# Print the best parameters found by GridSearchCV
cat("Best Parameters for Put Options from GridSearchCV:\n")
print(xgb_tune_put$bestTune)

# Extract the best parameters from GridSearchCV
best_params_put <- list(
  max_depth = xgb_tune_put$bestTune$max_depth,
  eta = xgb_tune_put$bestTune$eta,
  gamma = xgb_tune_put$bestTune$gamma,
  colsample_bytree = xgb_tune_put$bestTune$colsample_bytree,
  min_child_weight = xgb_tune_put$bestTune$min_child_weight,
  subsample = xgb_tune_put$bestTune$subsample,
  objective = "reg:gamma",               # Use gamma regression
  eval_metric = "rmse",
  base_score = mean(y_train_xgb_put),    # Initialize with mean target value
  seed = 1234                            # For reproducibility
)

# Train the model with the best parameters from GridSearchCV
xgb_model_grid_put <- xgb.train(
  params = best_params_put,
  data = dtrain_put,
  nrounds = 6250, # Use the best number of boosting rounds
  watchlist = list(train = dtrain_put, test = dtest_put),
  early_stopping_rounds = 10,             # Stop training if no improvement
  verbose = TRUE
)

# Make predictions and calculate RMSE and MAPE for GridSearchCV model
y_pred_grid_put <- predict(xgb_model_grid_put, dtest_put)

rmse_grid_put <- sqrt(mean((y_test_xgb_put - y_pred_grid_put)^2))
cat("RMSE with GridSearchCV parameters for Put options:", rmse_grid_put, "\n")

mape_grid_put <- mean(abs((y_pred_grid_put - y_test_xgb_put) / y_test_xgb_put)) * 100
cat("MAPE with GridSearchCV parameters for Put options:", mape_grid_put, "%\n")

# Prepare the data for saving
results <- data.frame(
  Date = test_data_put$Time,        # Assuming the 'Date' column exists in your testing data
  Actual = y_test_xgb_put,
  Predicted = y_pred_grid_put
)

# Save the results to an Excel file
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost GridSearchCV BTCPut.xlsx"
write.xlsx(results, output_file)

cat("Results saved to Excel file:", output_file, "\n")





# Load required libraries
library(readxl)

# File path
file_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results XGBoost GridSCV SPCall.xlsx"

# Read the Excel file
results_data <- read_excel(file_path)

# Ensure the columns 'Actual' and 'Predicted' are available
if (!all(c("Actual", "Predicted") %in% colnames(results_data))) {
  stop("Columns 'Actual' and 'Predicted' are not found in the file.")
}

# Extract Actual and Predicted values
actual <- results_data$Actual
predicted <- results_data$Predicted

# Calculate RMSE
rmse <- sqrt(mean((actual - predicted)^2))
cat("RMSE:", rmse, "\n")

# Calculate MAPE
mape <- mean(abs((actual - predicted) / actual)) * 100
cat("MAPE:", mape, "%\n")



##### Put dengan beda2 lag ####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Testing Data Put")

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  interval = 10,                    # Watchlist logging interval
  learning_rate = 0.02,             # This is the same as 'eta'
  objective = "reg:gamma",          # Use gamma loss
  max_bin = 256,                    # Adjust this based on your tree method
  max_delta_step = 0,               # Constraint for update step
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 6250,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,               # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term (note the corrected spelling)
  scale_pos_weight = 1,
  smooth_interval = 200,             # Smooth interval (not a standard parameter)
  subsample = 0.7,
  tree_method = "auto",              # Set to automatic
  tweedie_p = 1.5                   # Specify the Tweedie p value
)

# Define the lag configurations
lags <- list(
  lag1 = c("RVlag1", "BVlag1"),
  lag7 = c("RVlag7", "BVlag7", "SJVlag7"),
  lag15 = c("RVlag15", "BVlag15", "SJVlag15"),
  lag30 = c("RVlag30", "BVlag30", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_BTCPut_Lags.xlsx"
wb <- createWorkbook()

# Loop through each lag configuration
for (lag_name in names(lags)) {
  cat("\nTesting with lag:", lag_name, "\n")
  
  # Get the relevant lag features
  selected_features <- c("Asset Price", "Maturity", "Strike", lags[[lag_name]])
  
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
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Testing Data Put")

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  interval = 10,                    # Watchlist logging interval
  learning_rate = 0.02,             # This is the same as 'eta'
  objective = "reg:gamma",          # Use gamma loss
  max_bin = 256,                    # Adjust this based on your tree method
  max_delta_step = 0,               # Constraint for update step
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 6250,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,               # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term (note the corrected spelling)
  scale_pos_weight = 1,
  smooth_interval = 200,             # Smooth interval (not a standard parameter)
  subsample = 0.7,
  tree_method = "auto",              # Set to automatic
  tweedie_p = 1.5                   # Specify the Tweedie p value
)

# Define the feature configurations based on variations
variations <- list(
  RV = c("RVlag1", "RVlag7", "RVlag15", "RVlag30"),
  BV = c("BVlag1", "BVlag7", "BVlag15", "BVlag30"),
  SJV = c("SJVlag1","SJVlag7", "SJVlag15", "SJVlag30")
)

# Create a new workbook to save results
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_BTCPut_Variations.xlsx"
wb <- createWorkbook()

# Loop through each variation
for (var_name in names(variations)) {
  cat("\nTesting with variation:", var_name, "\n")
  
  # Get the relevant features for the current variation
  selected_features <- c("Asset Price", "Maturity", "Strike", variations[[var_name]])
  
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

##### Put dengan all data #####
# Load the necessary libraries
library(readxl)
library(dplyr)
library(xgboost)
library(openxlsx)

# Load the training and testing data
train_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Training Data Put")
test_data_put <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx", sheet = "Testing Data Put")

# Set up parameters for the XGBoost model
params <- list(
  base_margin_initialize = TRUE,
  colsample_bylevel = 1,
  colsample_bytree = 0.5,
  interval = 10,                    # Watchlist logging interval
  learning_rate = 0.02,             # This is the same as 'eta'
  objective = "reg:gamma",          # Use gamma loss
  max_bin = 256,                    # Adjust this based on your tree method
  max_delta_step = 0,               # Constraint for update step
  max_depth = 5,                    # Depth of the tree
  min_child_weight = 1,             # Minimum sum of instance weight
  min_split_loss = 0.01,            # Minimum loss reduction
  missing = -999,                   # Missing value marker
  n_estimators = 6250,              # Number of boosting rounds
  num_parallel_tree = 1,
  random_state = 1234,               # For reproducibility
  reg_alpha = 0,                     # L1 regularization term
  reg_lambda = 1,                    # L2 regularization term (note the corrected spelling)
  scale_pos_weight = 1,
  smooth_interval = 200,             # Smooth interval (not a standard parameter)
  subsample = 0.7,
  tree_method = "auto",              # Set to automatic
  tweedie_p = 1.5                   # Specify the Tweedie p value
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
output_file <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results XGBoost/Results_XGBoost_BTCPut_All.xlsx"
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
  selected_features <- c("Asset Price", "Maturity", "Strike", lags[[lag_name]])
  
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
