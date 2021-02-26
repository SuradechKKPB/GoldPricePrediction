#DAAN881 Group 3

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
df_gold_bank_change = read.csv("https://www.pontawee.com/wp-content/uploads/2021/02/changes_latest_as_at_february-2021_ifs.csv")
view(dfSummary(df_gold_bank_change))  #library(summarytools)
explore(df_gold_bank_change)  #library(explore)
makeDataReport(df_gold_bank_change,                
               file = "gold_bank_change.rmd",
               replace = TRUE)  #library(dataMaid)
DataExplorer::create_report(df_gold_bank_change)


