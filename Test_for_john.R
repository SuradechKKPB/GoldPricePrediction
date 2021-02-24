#Test for John

library(skimr)
library(DataExplorer)

df1 <- read.xlsx("C:/#Personal/Penn State/3. DAAN 881/Gold Market 5 Years.xlsx")

summary(df1)

skim(df1)

#Library package DataExplorer
#https://towardsdatascience.com/simple-fast-exploratory-data-analysis-in-r-with-dataexplorer-package-e055348d9619
#Creates an HTML file showing data exploration summaries
DataExplorer::create_report(df1)
