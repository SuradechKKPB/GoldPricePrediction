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
View(df_pontawee_summary_record)  #Viewing of some records has shown that some rows of GoldBarSellPrice and GoldBarBuyPrice do not have any values. We can separate this into three cases:
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
