#DAAN881 Group 3 26Feb2021

#Deliverable 3 Data preliminary analysis and investigating for issues

#Loading required libraries
library(summarytools)
library(explore)
library(dataMaid)
library(DataExplorer)
library(openxlsx)

#1st table buy records from PontaweeSales.xlsx file
df_pontawee_buy_record = read.xlsx(xlsxFile = "https://www.pontawee.com/wp-content/uploads/2021/02/PontaweeSales.xlsx", sheet = 1)
view(dfSummary(df_pontawee_buy_record))  #library(summarytools)
explore(df_pontawee_buy_record)  #library(explore)
makeDataReport(df_pontawee_buy_record,                
               file = "Pontawee_buy_record_EDA.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_pontawee_buy_record)

#2nd table sell records from PontaweeSales.xlsx file
df_pontawee_sell_record = read.xlsx(xlsxFile = "https://www.pontawee.com/wp-content/uploads/2021/02/PontaweeSales.xlsx", sheet = 2)
view(dfSummary(df_pontawee_sell_record))  #library(summarytools)
explore(df_pontawee_sell_record)  #library(explore)
makeDataReport(df_pontawee_sell_record,                
               file = "Pontawee_sell_record_EDA.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_pontawee_sell_record)

#3rd table summary records from PontaweeSales.xlsx file
df_pontawee_summary_record = read.xlsx(xlsxFile = "https://www.pontawee.com/wp-content/uploads/2021/02/PontaweeSales.xlsx", sheet = 3)
view(dfSummary(df_pontawee_summary_record))  #library(summarytools)
explore(df_pontawee_summary_record)  #library(explore)
makeDataReport(df_pontawee_summary_record,                
               file = "Pontawee_summary_record_EDA.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_pontawee_summary_record)

#4th table historical 5 years gold price data
df_gold_price = read.csv("https://www.pontawee.com/wp-content/uploads/2021/02/HistoricalQuotes.csv")
view(dfSummary(df_gold_price))  #library(summarytools)
explore(df_gold_price)  #library(explore)
makeDataReport(df_gold_price,                
               file = "Gold_price_EDA.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_gold_price)

#5th table Historical major fund gold holding volume from 2003
df_fund_holding = read.xlsx(xlsxFile = "https://www.pontawee.com/wp-content/uploads/2021/02/2021-January-ETF-Holdings-Flows.xlsx", sheet = "All holdings by day", startRow = 6)
view(dfSummary(df_fund_holding))  #library(summarytools)
explore(df_fund_holding)  #library(explore)
makeDataReport(df_fund_holding,                
               file = "Fund_holding.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_fund_holding)

#6th table Historical major fund flow from 2003
df_fund_flow = read.xlsx(xlsxFile = "https://www.pontawee.com/wp-content/uploads/2021/02/2021-January-ETF-Holdings-Flows.xlsx", sheet = "All flows US$ by day", startRow = 6)
view(dfSummary(df_fund_flow))  #library(summarytools)
explore(df_fund_flow)  #library(explore)
makeDataReport(df_fund_flow,                
               file = "Fund_flow.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_fund_flow)

#7th table Gold supply and demand statistics
#No EDA performed on this table

#8th table Changes in Gold held by central banks
df_gold_bank_change = read.xlsx(xlsxFile = "https://pennstateoffice365-my.sharepoint.com/:x:/g/personal/wcb8_psu_edu/EXKIuXjDBeBHtBp76wqkOacBUuvvFE8rLRzRviDMx_D5_A?e=M972QE", sheet = 1)
view(dfSummary(df_gold_bank_change))  #library(summarytools)
explore(df_gold_bank_change)  #library(explore)
makeDataReport(df_gold_bank_change,                
               file = "gold_bank_change.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_gold_bank_change)

#Deliverable 4 Data Cleaning and plan for data transformation

#Decide to drop the 1st table (buy records) and 2nd table (sell records) from PontaweeSales.xlsx file from the analysis as there are a lot of inconsistencies with the 3rd table (Summary buy and sell transactions) and, as confirmed with the owners, the only way to make a good data cleaning is by re-key from hard copies.

#Take a look at Table 3 Summary buy and sell transactions
library(skimr)
library(openxlsx)
library(dplyr)
library(lubridate)
library(data.table)
library(tidyr)

df_pontawee_summary_record = read.xlsx(xlsxFile = "https://www.pontawee.com/wp-content/uploads/2021/02/PontaweeSales.xlsx", sheet = 3)
skim(df_pontawee_summary_record)
glimpse(df_pontawee_summary_record) #We observed that there are no real date-time column.  Month column is read as character so it should be converted to factor.  Year and date columns are read as double so it should be converted to integer.  There are a lot of missing values in Reset_stock_date so "NA" or "No" values should be filled-in for the missing cells.

df_pontawee_summary_record$new_date <- as.Date(format(dmy(paste(df_pontawee_summary_record$Date, df_pontawee_summary_record$Month, df_pontawee_summary_record$Year)), '%d/%m/%Y'), '%d/%m/%Y')

df_pontawee_summary_record$Month <- as.factor(df_pontawee_summary_record$Month)
df_pontawee_summary_record$Year <- as.integer(df_pontawee_summary_record$Year)
df_pontawee_summary_record$Date <- as.integer(df_pontawee_summary_record$Date)
df_pontawee_summary_record$Reset_Stock_date[is.na(df_pontawee_summary_record$Reset_Stock_date)] = "No"
df_pontawee_summary_record$Reset_Stock_date <- as.factor(df_pontawee_summary_record$Reset_Stock_date)

df_pontawee_summary_record$Sell_weight[is.na(df_pontawee_summary_record$Sell_weight)] = 0 #Replace missing values with 0
df_pontawee_summary_record$Buy_weight[is.na(df_pontawee_summary_record$Buy_weight)] = 0 #Replace missing values with 0

glimpse(df_pontawee_summary_record) #The columns are converted to correct data type now.
View(df_pontawee_summary_record)  #Viewing of some records has shown that some rows of GoldBarSellPrice and GoldBarBuyPrice do not have any values. We can separate this into two main cases:
# 1. Have values on GoldBarSellPrice but have missing values on GoldBarBuy price or vice versa, the missing values can be calculated as per typical margin on buy and sell price is usually 100 baht
# 2. No values on both cells.  We decide to replace them manually due to small number of missing cells.

setDT(df_pontawee_summary_record) # <- convert to data.table
# going column-by-column, count NA
df_pontawee_summary_record[ , lapply(.SD, function(x) sum(is.na(x))), by = Year]
skim(df_pontawee_summary_record)
#Case 1: Missing value on GoldBarSellPrice but has value on GoldBarBuyPrice
#View the records
df_pontawee_summary_record[is.na(df_pontawee_summary_record$GoldBarSellPrice) & !is.na(df_pontawee_summary_record$GoldBarBuyPrice)]

#Change the missing values by calculation
df_pontawee_summary_record$GoldBarSellPrice[is.na(df_pontawee_summary_record$GoldBarSellPrice) & !is.na(df_pontawee_summary_record$GoldBarBuyPrice)] = df_pontawee_summary_record$GoldBarBuyPrice[is.na(df_pontawee_summary_record$GoldBarSellPrice) & !is.na(df_pontawee_summary_record$GoldBarBuyPrice)] + 100

skim(df_pontawee_summary_record)

#Case 2: Missing value on GoldBarBuyPrice but has value on GoldBarSellPrice
#View the records
df_pontawee_summary_record[!is.na(df_pontawee_summary_record$GoldBarSellPrice) & is.na(df_pontawee_summary_record$GoldBarBuyPrice)]

#Change the missing values by calculation
df_pontawee_summary_record$GoldBarBuyPrice[!is.na(df_pontawee_summary_record$GoldBarSellPrice) & is.na(df_pontawee_summary_record$GoldBarBuyPrice)] = df_pontawee_summary_record$GoldBarSellPrice[!is.na(df_pontawee_summary_record$GoldBarSellPrice) & is.na(df_pontawee_summary_record$GoldBarBuyPrice)] - 100

skim(df_pontawee_summary_record)

#There are still 48 missing values
df_pontawee_summary_record[is.na(df_pontawee_summary_record$GoldBarSellPrice) & is.na(df_pontawee_summary_record$GoldBarBuyPrice)]

#As these missing values are not contributed to any business questions(it was missing because there is no buy or sell transaction on that day), we'll just impute from previous price here.
df_pontawee_summary_record = df_pontawee_summary_record %>%
  fill(GoldBarSellPrice) %>%
  fill(GoldBarBuyPrice)

skim(df_pontawee_summary_record)

#We now have a cleaned data frame for Table 3.

#Deliverable 5

#It is found that there are no data cleanliness issues on the remaining tables.  Although there are some outliers, they are the real data points.  Therefore, we'll just leave them as per current stage and continue with modeling steps.

#First, let's visualize the gold price in THB from the cleaned table that we progressed so far.
library(ggplot2)
ggplot(data=df_pontawee_summary_record, aes(x=new_date, y=GoldBarSellPrice)) +
  geom_line() 

#It is observed that there are still some records that show gold price = 0.  We'll continue cleaning the data by assigning them to 0 and impute with previous price
df_pontawee_summary_record <- df_pontawee_summary_record %>% 
  mutate(GoldBarSellPrice = replace(GoldBarSellPrice, GoldBarSellPrice=="0", NA)) %>%
  mutate(GoldBarBuyPrice = replace(GoldBarBuyPrice, GoldBarBuyPrice=="0", NA))
df_pontawee_summary_record = df_pontawee_summary_record %>%
  fill(GoldBarSellPrice) %>%
  fill(GoldBarBuyPrice)

ggplot(data=df_pontawee_summary_record, aes(x=new_date, y=GoldBarSellPrice)) +
  geom_line() 

#We have a good dataframe now.  Next, let's try to visualize the gold price in THB versus USD.

#Convert Date columne to be a date time object
df_gold_price$Date <- as.Date(df_gold_price$Date, "%m/%d/%Y")
df_gold_price <- df_gold_price %>%
  filter(Date >='2019-01-01') %>%
  arrange(Date)
df_gold_price
names(df_gold_price)[1] <- "new_date"  #Rename column to be the same for both data frame
glimpse(df_pontawee_summary_record)

library(dplyr)
df_combined <- left_join(df_pontawee_summary_record, df_gold_price, by="new_date")
df_combined[1:20]
library("tseries")    
range01 <- function(x){(x-min(x, na.rm = T))/(max(x, na.rm = T)-min(x, na.rm = T))}
comb_ts <- cbind(range01(df_combined$GoldBarSellPrice), range01(df_combined$Open))# please make sure the length of both your timeseries
comb_ts
plot.ts(comb_ts, plot.type = "single", main = "Gold price comparison between THB and USD")
#It appears that there are some slight variation between the gold price in USD and in THB due to changes in currency exchange rate.  We will focus on the price in THB here to answer business's question.

#First, let's split the data into training and test set.  As we have full data for 2019 and 2020 and almost one month data in 2021.  To get sufficient testing data, we've selected the last three months data as test set.

df_train <- df_pontawee_summary_record %>%
  filter(new_date <'2020-10-01')
df_test <- df_pontawee_summary_record %>%
  filter(new_date >='2020-10-01')
length(df_train$GoldBarSellPrice)
length(df_test$GoldBarSellPrice)
length(df_pontawee_summary_record$GoldBarSellPrice)

#Let's try a univariate time series model for gold price in THB. 

#ARIMA modeling
library(astsa)
library(forecast)
acf2(df_train$GoldBarSellPrice) #ACF plot shows that the data is not stationary, we'll try first differencing here.
diff1 = diff(df_train$GoldBarSellPrice,1)
acf2(diff1)
plot(diff1, type = 'l')

sarima(df_train$GoldBarSellPrice, p = 0,d = 1,q = 0)
sarima(df_train$GoldBarSellPrice, p = 0,d = 1,q = 0,P = 2,D = 0,Q = 0, S = 7)
sarima(df_train$GoldBarSellPrice, p = 0,d = 1,q = 0,P = 0,D = 0,Q = 2, S = 7)
sarima(df_train$GoldBarSellPrice, p = 0,d = 1,q = 0,P = 1,D = 0,Q = 2, S = 7)
sarima(df_train$GoldBarSellPrice, p = 0,d = 1,q = 0,P = 2,D = 0,Q = 2, S = 7)
sarima(df_train$GoldBarSellPrice, p = 0,d = 1,q = 0,P = 2,D = 0,Q = 1, S = 7)

model1 <- auto.arima(df_train$GoldBarSellPrice, trace = T) # We got the best fit model is ARIMA(0,1,0) - this is a white noise model indicating that the forecast will rely heavily on its mean or intercept parameter
model1
model2 <- arima(df_train$GoldBarSellPrice, order = c(0,1,0), seasonal = list(order = c(2,0,0), period = 7))
model2

acf2(diff1)
predict1 <- forecast(model1, 115)
#For plotting
autoplot(predict1) + autolayer(predict1$fitted) + autolayer(predict1$mean) + geom_line(aes((length(df_train$GoldBarSellPrice)+1):length(df_pontawee_summary_record$GoldBarSellPrice), df_test$GoldBarSellPrice, group = 1), color = "red")  

autoplot(predict2) + autolayer(predict2$fitted) + autolayer(predict2$mean) + geom_line(aes((length(df_train$GoldBarSellPrice)+1):length(df_pontawee_summary_record$GoldBarSellPrice), df_test$GoldBarSellPrice, group = 1), color = "red") 

library(Metrics)
Metrics::rmse(df_train$GoldBarSellPrice, predict1$fitted)
Metrics::rmse(df_train$GoldBarSellPrice, predict2$fitted)
Metrics::rmse(df_test$GoldBarSellPrice, predict1$mean)
Metrics::rmse(df_test$GoldBarSellPrice, predict2$mean)

#Forecasting with snaive
predict3 <- snaive(ts(df_train$GoldBarSellPrice, frequency = 7), h = 115)
autoplot(predict3) + autolayer(predict3$fitted) + autolayer(predict3$mean) + geom_line(aes((length(df_train$GoldBarSellPrice)+1):length(df_pontawee_summary_record$GoldBarSellPrice)/7, df_test$GoldBarSellPrice, group = 1), color = "red") 
Metrics::rmse(tail(df_train$GoldBarSellPrice, -7), tail(predict3$fitted, -7)) #Remove first 7 values as seasonal naive prediction will produce NA
Metrics::rmse(df_test$GoldBarSellPrice, predict3$mean)

# Fitting with tbats
tbats_model <- tbats(df_train$GoldBarSellPrice, seasonal.periods = 7)
tbats_model

# forecast for next months
tbats_forecast <- forecast(tbats_model, h=115)
print(tbats_forecast)
# Plotting the model
autoplot(tbats_forecast) + autolayer(tbats_forecast$fitted) + autolayer(tbats_forecast$mean) + geom_line(aes((length(df_train$GoldBarSellPrice)+1):length(df_pontawee_summary_record$GoldBarSellPrice), df_test$GoldBarSellPrice, group = 1), color = "red") 
Metrics::rmse(df_train$GoldBarSellPrice, tbats_forecast$fitted)
Metrics::rmse(df_test$GoldBarSellPrice, tbats_forecast$mean)


library(prophet)
df <- data.frame(df_train$new_date, df_train$GoldBarSellPrice)
head(df)
df <- df %>% rename(ds = df_train.new_date, y = df_train.GoldBarSellPrice)


prophet_clean_data <- prophet(df, yearly.seasonality = T)
prophet_clean_data

# forecast for next month
future <- make_future_dataframe(prophet_clean_data, periods = 115)
prophet_clean_data_forecast <- predict(prophet_clean_data,future)
print(prophet_clean_data_forecast)
# Plotting the model
aaa <- plot(prophet_clean_data, prophet_clean_data_forecast, main = 'Prophet')
aaa
tail(prophet_clean_data_forecast$yhat, n = 115)
#Store the forecast data
pred <- tail(prophet_clean_data_forecast$yhat, n = 115)

# Plotting the model
plot(df_pontawee_summary_record$new_date, df_pontawee_summary_record$GoldBarSellPrice, type = 'l')
#plotting the model fitting value of train set
lines(x = df_train$new_date,y = df_train$GoldBarSellPrice, col = 2)
abline(v = df_train$new_date[639])


#plotting the model fitting value of test set
lines(x = df_test$new_date, y = predict1$mean, col = 'blue')

#plot(tbats_forecast)
lines(x = df_train$new_date,y = tbats_forecast$fitted, col = 2)
# Add a line with actual values
lines(x = df_test$new_date, y = tbats_forecast$mean, col = 'blue')

lines(x = df_pontawee_summary_record$new_date,y = prophet_clean_data_forecast$yhat, col = 3)

write.csv(df_pontawee_summary_record, "C:/Users/szjt/OneDrive - Chevron/DAAN881/df_summary.csv")

m <- prophet(df, yearly.seasonality = T)
future <- make_future_dataframe(m, periods = 115)
forecast <- predict(m, future)
plot(m, forecast, main = "Gold price forecast from prophet model")

prophet_plot_components(m, forecast)

Metrics::rmse(df_train$GoldBarSellPrice, prophet_clean_data_forecast$yhat[1:639])
Metrics::rmse(df_test$GoldBarSellPrice, prophet_clean_data_forecast$yhat[640:754])

