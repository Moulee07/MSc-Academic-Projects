library(dplyr)
library(forecast)
library(prophet)
library(ggplot2)
library(readxl)
library(tidyverse)
library(openxlsx)

## Read the data
data <-read_excel('/Users/mahendra/Downloads/Sales_Data_Final (1).xlsx')
df1 <- read_excel('/Users/mahendra/Downloads/Sales_Data_Final (1).xlsx')
# Values are seperated by comma hence replacing the comma(,)
df1$acf_new = gsub(',','',df1$ACF_New)
df1$acf_new = as.numeric(df1$acf_new)
# Converting the date(String) to strptime which is further converted to date format
df1$cashflowdate = strptime(df1$CashflowDate,format = "%Y-%m-%d")
df1$cashflowdate = as.Date(df1$cashflowdate)
max<- as.Date('2024-05-28')
min_date = as.Date('2022-01-01')
all_dates <- data.frame(cashflowdate = seq(min_date, max, by = "day"))
colnames(all_dates)[1]="ds" 



# Extract year and month from the cashflowdate
df1 <- df1 %>% filter(cashflowdate<='2024-05-28')
df1 <- df1 %>%
  mutate(
    Year = year(cashflowdate),
    Month = month(cashflowdate),
    Week = week(cashflowdate),
    Day = weekdays.Date(cashflowdate),
    days = day(cashflowdate)
  )

# holiday_flag initiation
df1 <- df1 %>%
  mutate(holiday_flag = ifelse(Month == 12 & days == 25,1,0))%>%
  mutate(holiday_flag = ifelse(ACF_New==0,1,0))

#christmas_flag initiation
df1 <- df1 %>%
  mutate(christmas_flag = ifelse(holiday_flag == 1 & Month == 12 & days == 25,1,0)) %>%
  mutate(holiday_flag = ifelse(christmas_flag == 1, 0, holiday_flag))


df1 <- df1 %>%
  mutate(subliquidityitem = ifelse(SubLiquidityItem == '','unknown',SubLiquidityItem))


# Adding a unique row identifier to ensure no data is aggregated unintentionally
df1 <- df1 %>%
  mutate(row_id = row_number())



# Performing one-hot encoding for subliquidityitem
df1_encoded <- df1 %>%
  pivot_wider(names_prefix = "channel_",
              names_from = subliquidityitem,
              values_from = subliquidityitem,
              values_fill = list(subliquidityitem = 0), # Fill absent factors with 0
              values_fn = list(subliquidityitem = length)) %>%
  mutate(across(starts_with("channel_"), ~ as.integer(. > 0))) %>%
  select(-row_id) # Remove row identifier if no longer needed


# Data narrowing by selecting the required columns
df1_encoded <- df1_encoded %>%
  select(cashflowdate,
         acf_new,
         Country,
         holiday_flag,
         christmas_flag,
         starts_with("channel")
  )

# Renaming the columns as required by prophet model
colnames(df1_encoded)[1]="ds"        
colnames(df1_encoded)[2]="y"


# Summing y by cashflowdate and holiday_flag
group_columns <- df1_encoded %>%
  select(ds, Country,starts_with("channel_")) %>%
  colnames()

# Then, use the selected columns in group_by
df1_summed <- df1_encoded %>%
  group_by(across(all_of(group_columns))) %>%
  summarise(y = sum(y, na.rm = TRUE), .groups = 'drop')


# Initialize workbook
all_results <- data.frame()
################################ (Demand Forecasting,Retail)
forecast_country <- function(country_name) {
  # filter the channel 1 to forecast for channel 1 
  AT_df_encoded_channel1 <- df1_summed %>% filter(Country == country_name) %>% filter(channel_Channel1 == 1)
  # Rechecking by selecting only the required columns
  AT_df_encoded_channel1 <- AT_df_encoded_channel1 %>% select(ds, y, Country, starts_with("channel_"), ends_with("_flag"))
  # If a date being missed in the sales_data it is observed to be a holiday or there by joining the dates
  AT_df_encoded_channel1_complete <- all_dates %>% left_join(AT_df_encoded_channel1, by = 'ds')
  # Rechecking christmas and holiday flags
  AT_df_encoded_channel1_complete <- AT_df_encoded_channel1_complete %>%
    mutate(holiday_flag = ifelse(y == 0, 1, 0)) %>%
    mutate(christmas_flag = ifelse(holiday_flag == 1 &  month(ds)== c(12) & day(ds) %in% c(25, 1), 1, 0))
  # Handling Null values
  AT_df_encoded_channel1_complete <- AT_df_encoded_channel1_complete %>%
    mutate(y = ifelse(is.na(y), 0, y)) %>%
    mutate(christmas_flag = ifelse(is.na(christmas_flag), 0, christmas_flag)) %>%
    mutate(holiday_flag = ifelse(is.na(holiday_flag), 0, holiday_flag)) %>%
    mutate(channel_Channel1 = ifelse(is.na(channel_Channel1), 0, channel_Channel1))%>%
    mutate(channel_Channel2 = ifelse(is.na(channel_Channel2), 0, channel_Channel2))%>%
    mutate(channel_Channel3 = ifelse(is.na(channel_Channel3), 0, channel_Channel3))%>%
    mutate(channel_Channel4 = ifelse(is.na(channel_Channel4), 0, channel_Channel4))%>%
    mutate(Country = ifelse(is.na(Country),country_name,Country))
  # Rechecking christmas and holiday flags
  AT_df_encoded_channel1_complete <- AT_df_encoded_channel1_complete %>%
    mutate(holiday_flag = ifelse(y == 0, 1, 0)) %>%
    mutate(christmas_flag = ifelse(holiday_flag == 1 &  month(ds)== c(12) & day(ds) %in% c(25, 1), 1, 0))
  
  ### Holidays check
  holidays_dates = subset(AT_df_encoded_channel1_complete,holiday_flag==1)
  holidays_dates = holidays_dates$ds
  holidays = tibble(holiday='holiday',
                    ds=unique(holidays_dates),
                    lower_window=-4,
                    upper_window =+2)
  # Creating Grid for prophet for performing cross-validation
  prophet_grid <- expand.grid(changepoint_prior_scale = c(0.05, 0.1,0.5),
                              seasonality_prior_scale = c(5, 10,15), 
                              holidays_prior_scale = c(5, 10),
                              seasonality.mode = c('multiplicative', 'additive'))
  
  results <- vector(mode = 'numeric', length = nrow(prophet_grid))
  names(results) <- paste0("Model_", 1:nrow(prophet_grid))
  
  # performing grid search for different combinations of given grid search parameters for hyperparameter tuning
  for (i in 1:nrow(prophet_grid)) {
    try({
      parameters <- prophet_grid[i, ]
      
      m <- prophet(yearly.seasonality = TRUE,
                   weekly.seasonality = TRUE,
                   daily.seasonality = FALSE,
                   holidays = holidays,
                   seasonality.mode = parameters$seasonality.mode,
                   seasonality.prior.scale = parameters$seasonality_prior_scale,
                   holidays.prior.scale = parameters$holidays_prior_scale,
                   changepoint.prior.scale = parameters$changepoint_prior_scale)
      m <- add_regressor(m, "christmas_flag")

      m <- fit.prophet(m, AT_df_encoded_channel1_complete)
      
      df.cv <- cross_validation(model = m,
                                horizon = 90,
                                units = "days",
                                period = 7,
                                initial = 700
                                )
      
      df.perf <- performance_metrics(df.cv, metrics = c('mae', 'mse'))
      results[i] <- df.perf$mae
      
    }, silent = TRUE)
  }
  
  prophet_grid <- cbind(prophet_grid, results)
  # out of the generated values from prophe grid search select the parameters with least error value for hyper parameter tuning
  best_params <- prophet_grid[prophet_grid$results == min(results), ]
  # If multiple parameters have same error value then select based on rank
  if (length(best_params) != 1) {
    best_model_ranked <- best_params %>%
      arrange(changepoint_prior_scale, seasonality_prior_scale, holidays_prior_scale)
    
    best_params <- best_model_ranked[1, ]
  }
  # Initialize final results data frame
  final_results <- data.frame()
  i=0
  # Iterative Rolling Forecast
  for (start_date in seq(as.Date('2023-12-06'), by = "week", length.out = 13)) {
    # train and test split
    training <- AT_df_encoded_channel1_complete %>% filter(ds <= start_date) %>%
      select(ds, y, christmas_flag)
    test <- AT_df_encoded_channel1_complete %>% filter(ds > start_date & ds <= start_date + 90) %>%
      select(ds, y, christmas_flag)
    
    m <- prophet(holidays = holidays,
                 yearly.seasonality = TRUE,
                 weekly.seasonality = TRUE,
                 daily.seasonality = FALSE,
                 seasonality.mode = best_params$seasonality.mode,
                 seasonality.prior.scale = best_params$seasonality_prior_scale,
                 holidays.prior.scale = best_params$holidays_prior_scale,
                 changepoint.prior.scale = best_params$changepoint_prior_scale)
    m <- add_regressor(m, "christmas_flag")

    m <- fit.prophet(m, training)
    
    # future data_frame to accomodate the forecast values
    future <- make_future_dataframe(m, periods = nrow(test))
    
    future <- future %>%
      left_join(AT_df_encoded_channel1_complete %>% select(ds, christmas_flag), by = "ds")
    
    future <- future %>%
      mutate(across(starts_with("christmas"), ~ ifelse(is.na(.), 0, .)))
    
    forecast <- predict(m, future)
    ## check handling if the forecast is predicting any non-zero values in weekends
    forecast <- forecast %>%
      mutate(
        weekend_flag = ifelse(wday(ds) %in% c(1, 7), 1, 0),
        yhat = ifelse(weekend_flag == 1, 0, yhat)
      )
    
    predictions <- tail(forecast$yhat, nrow(test))
    actuals <- test$y
    
    delta <- predictions - actuals
    percentage_error <- (predictions / actuals - 1)
    percentage_delta <- percentage_error * 100
    accuracy <- 1 - percentage_error
    accuracy_percentage <- accuracy * 100
    mahe = seq(as.Date('2023-12-06'), by = "week", length.out = 13)
    i=i+1
    tryCatch({
      result_df <- data.frame(
        country = country_name,
        AsOfDate = mahe[i],
        AsOfWeek = week(start_date) - week(test$ds),
        IsoWeek = week(test$ds),
        Actuals = actuals,
        Predictions = ifelse(actuals==0,0,predictions),
        cashflowDate = test$ds,
        delta = ifelse(actuals==0,0,delta),
        per_delta_value = ifelse(actuals==0,0,percentage_error),
        percentage_delta_value = ifelse(actuals==0,0,percentage_delta),
        accuracy = ifelse(actuals==0,0,accuracy_percentage)
      )
      all_results <<- bind_rows(all_results, result_df)
      print(result_df)
      
    }, error = function(e) {
      cat("Error occurred while creating or writing result_df for country:", country_name, " and start_date: ", start_date, "\n")
      cat("Error message:", e$message, "\n")
    })
  }
  # Append final_results to the global all_results dataframe
  #all_results <<- bind_rows(all_results, final_results)
  # Plot the forecast with test data
  # Plot the forecast with test data
  forecast_plot <- ggplot() +
    geom_line(data = forecast, aes(x = as.POSIXct(ds), y = yhat, color = "Forecast")) +
    geom_ribbon(data = forecast, aes(x = as.POSIXct(ds), ymin = yhat_lower, ymax = yhat_upper), alpha = 0.2, fill = "blue") +
    geom_point(data = training, aes(x = as.POSIXct(ds), y = y, color = "Training Data")) +
    geom_point(data = test, aes(x = as.POSIXct(ds), y = y, color = "Test Data")) +
    labs(title = paste("Prophet Forecast for", country_name),
         x = "Date", y = "ACF New") +
    scale_color_manual(values = c("Forecast" = "skyblue", "Training Data" = "black","Test Data"="red"))
  
  # Save the forecast plot
  ggsave(paste0("/Users/mahendra/Downloads/forecasts_cv/forecast_", country_name, ".png"), plot = forecast_plot, width = 10, height = 6)
  
  
  # Ensure the test data is in POSIXct format
  # Ensure the test data is in POSIXct format
  test$ds <- as.POSIXct(test$ds)

  }

# After calling forecast_country for all countries, save all_results to an Excel file
save_results_to_excel <- function() {
  excel_file <- "/Users/mahendra/Downloads/forecasts_cv/forecast_channel_1.xlsx"
  sheet_name <- "ForecastResults"
  
  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet = sheet_name, all_results)
  saveWorkbook(wb, file = excel_file, overwrite = TRUE)
}
unique_countries <- unique(df1$Country)
lapply(unique_countries, forecast_country)

save_results_to_excel()

## Channel 2

# Initialize workbook
all_results <- data.frame()
################################ (Demand Forecasting,Retail)
forecast_country <- function(country_name) {
  # filter the channel 1 to forecast for channel 1 
  AT_df_encoded_channel1 <- df1_summed %>% filter(Country == country_name) %>% filter(channel_Channel2 == 1)
  # Rechecking by selecting only the required columns
  AT_df_encoded_channel1 <- AT_df_encoded_channel1 %>% select(ds, y, Country, starts_with("channel_"), ends_with("_flag"))
  # If a date being missed in the sales_data it is observed to be a holiday or there by joining the dates
  AT_df_encoded_channel1_complete <- all_dates %>% left_join(AT_df_encoded_channel1, by = 'ds')
  # Rechecking christmas and holiday flags
  AT_df_encoded_channel1_complete <- AT_df_encoded_channel1_complete %>%
    mutate(holiday_flag = ifelse(y == 0, 1, 0)) %>%
    mutate(christmas_flag = ifelse(holiday_flag == 1 &  month(ds)== c(12) & day(ds) %in% c(25, 1), 1, 0))
  # Handling Null values
  AT_df_encoded_channel1_complete <- AT_df_encoded_channel1_complete %>%
    mutate(y = ifelse(is.na(y), 0, y)) %>%
    mutate(christmas_flag = ifelse(is.na(christmas_flag), 0, christmas_flag)) %>%
    mutate(holiday_flag = ifelse(is.na(holiday_flag), 0, holiday_flag)) %>%
    mutate(channel_Channel1 = ifelse(is.na(channel_Channel1), 0, channel_Channel1))%>%
    mutate(channel_Channel2 = ifelse(is.na(channel_Channel2), 0, channel_Channel2))%>%
    mutate(channel_Channel3 = ifelse(is.na(channel_Channel3), 0, channel_Channel3))%>%
    mutate(channel_Channel4 = ifelse(is.na(channel_Channel4), 0, channel_Channel4))%>%
    mutate(Country = ifelse(is.na(Country),country_name,Country))
  # Rechecking christmas and holiday flags
  AT_df_encoded_channel1_complete <- AT_df_encoded_channel1_complete %>%
    mutate(holiday_flag = ifelse(y == 0, 1, 0)) %>%
    mutate(christmas_flag = ifelse(holiday_flag == 1 &  month(ds)== c(12) & day(ds) %in% c(25, 1), 1, 0))
  
  ### Holidays check
  holidays_dates = subset(AT_df_encoded_channel1_complete,holiday_flag==1)
  holidays_dates = holidays_dates$ds
  holidays = tibble(holiday='holiday',
                    ds=unique(holidays_dates),
                    lower_window=-4,
                    upper_window =+2)
  # Creating Grid for prophet for performing cross-validation
  prophet_grid <- expand.grid(changepoint_prior_scale = c(0.05, 0.1,0.5),
                              seasonality_prior_scale = c(5, 10,15), 
                              holidays_prior_scale = c(5, 10),
                              seasonality.mode = c('multiplicative', 'additive'))
  
  results <- vector(mode = 'numeric', length = nrow(prophet_grid))
  names(results) <- paste0("Model_", 1:nrow(prophet_grid))
  
  # performing grid search for different combinations of given grid search parameters for hyperparameter tuning
  for (i in 1:nrow(prophet_grid)) {
    try({
      parameters <- prophet_grid[i, ]
      
      m <- prophet(yearly.seasonality = TRUE,
                   weekly.seasonality = TRUE,
                   daily.seasonality = FALSE,
                   holidays = holidays,
                   seasonality.mode = parameters$seasonality.mode,
                   seasonality.prior.scale = parameters$seasonality_prior_scale,
                   holidays.prior.scale = parameters$holidays_prior_scale,
                   changepoint.prior.scale = parameters$changepoint_prior_scale)
      m <- add_regressor(m, "christmas_flag")
      
      m <- fit.prophet(m, AT_df_encoded_channel1_complete)
      
      df.cv <- cross_validation(model = m,
                                horizon = 90,
                                units = "days",
                                period = 7,
                                initial = 700
      )
      
      df.perf <- performance_metrics(df.cv, metrics = c('mae', 'mse'))
      results[i] <- df.perf$mae
      
    }, silent = TRUE)
  }
  
  prophet_grid <- cbind(prophet_grid, results)
  # out of the generated values from prophe grid search select the parameters with least error value for hyper parameter tuning
  best_params <- prophet_grid[prophet_grid$results == min(results), ]
  # If multiple parameters have same error value then select based on rank
  if (length(best_params) != 1) {
    best_model_ranked <- best_params %>%
      arrange(changepoint_prior_scale, seasonality_prior_scale, holidays_prior_scale)
    
    best_params <- best_model_ranked[1, ]
  }
  # Initialize final results data frame
  final_results <- data.frame()
  i=0
  # Iterative Rolling Forecast
  for (start_date in seq(as.Date('2023-12-06'), by = "week", length.out = 13)) {
    # train and test split
    training <- AT_df_encoded_channel1_complete %>% filter(ds <= start_date) %>%
      select(ds, y, christmas_flag)
    test <- AT_df_encoded_channel1_complete %>% filter(ds > start_date & ds <= start_date + 90) %>%
      select(ds, y, christmas_flag)
    
    m <- prophet(holidays = holidays,
                 yearly.seasonality = TRUE,
                 weekly.seasonality = TRUE,
                 daily.seasonality = FALSE,
                 seasonality.mode = best_params$seasonality.mode,
                 seasonality.prior.scale = best_params$seasonality_prior_scale,
                 holidays.prior.scale = best_params$holidays_prior_scale,
                 changepoint.prior.scale = best_params$changepoint_prior_scale)
    m <- add_regressor(m, "christmas_flag")
    
    m <- fit.prophet(m, training)
    
    # future data_frame to accomodate the forecast values
    future <- make_future_dataframe(m, periods = nrow(test))
    
    future <- future %>%
      left_join(AT_df_encoded_channel1_complete %>% select(ds, christmas_flag), by = "ds")
    
    future <- future %>%
      mutate(across(starts_with("christmas"), ~ ifelse(is.na(.), 0, .)))
    
    forecast <- predict(m, future)
    ## check handling if the forecast is predicting any non-zero values in weekends
    forecast <- forecast %>%
      mutate(
        weekend_flag = ifelse(wday(ds) %in% c(1, 7), 1, 0),
        yhat = ifelse(weekend_flag == 1, 0, yhat)
      )
    
    predictions <- tail(forecast$yhat, nrow(test))
    actuals <- test$y
    # Accuarcy Check
    delta <- predictions - actuals
    percentage_error <- (predictions / actuals - 1)
    percentage_delta <- percentage_error * 100
    accuracy <- 1 - percentage_error
    accuracy_percentage <- accuracy * 100
    # generatig sequence of dates for AsOfDate generation, will be updated for every iteration from the list of dates available
    mahe = seq(as.Date('2023-12-06'), by = "week", length.out = 13)
    i=i+1
    tryCatch({
      result_df <- data.frame(
        country = country_name,
        AsOfDate = mahe[i],
        AsOfWeek = week(start_date) - week(test$ds),
        IsoWeek = week(test$ds),
        Actuals = actuals,
        Predictions = ifelse(actuals==0,0,predictions),
        cashflowDate = test$ds,
        delta = ifelse(actuals==0,0,delta),
        per_delta_value = ifelse(actuals==0,0,percentage_error),
        percentage_delta_value = ifelse(actuals==0,0,percentage_delta),
        accuracy = ifelse(actuals==0,0,accuracy_percentage)
      )
      all_results <<- bind_rows(all_results, result_df)
      print(result_df)
      
    }, error = function(e) {
      cat("Error occurred while creating or writing result_df for country:", country_name, " and start_date: ", start_date, "\n")
      cat("Error message:", e$message, "\n")
    })
  }
  # Append final_results to the global all_results dataframe
  # Plot the forecast with test data
  # Plot the forecast with test data
  forecast_plot <- ggplot() +
    geom_line(data = forecast, aes(x = as.POSIXct(ds), y = yhat, color = "Forecast")) +
    geom_ribbon(data = forecast, aes(x = as.POSIXct(ds), ymin = yhat_lower, ymax = yhat_upper), alpha = 0.2, fill = "blue") +
    geom_point(data = training, aes(x = as.POSIXct(ds), y = y, color = "Training Data")) +
    geom_point(data = test, aes(x = as.POSIXct(ds), y = y, color = "Test Data")) +
    labs(title = paste("Prophet Forecast for", country_name),
         x = "Date", y = "ACF New") +
    scale_color_manual(values = c("Forecast" = "skyblue", "Training Data" = "black","Test Data"="red"))
  
  # Save the forecast plot
  ggsave(paste0("/Users/mahendra/Downloads/forecasts_cv/forecast_", country_name, ".png"), plot = forecast_plot, width = 10, height = 6)
  
  
}

# After calling forecast_country for all countries, save all_results to an Excel file
save_results_to_excel <- function() {
  excel_file <- "/Users/mahendra/Downloads/forecasts_cv/forecast_channel_2.xlsx"
  sheet_name <- "ForecastResults"
  
  wb <- createWorkbook()
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet = sheet_name, all_results)
  saveWorkbook(wb, file = excel_file, overwrite = TRUE)
}
# taking list of unique countries
unique_countries <- unique(df1$Country)
# apply to the function to iterate over the distinct countries.
lapply(unique_countries, forecast_country)

save_results_to_excel()



