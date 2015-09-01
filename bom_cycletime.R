setwd("~/RProgramming/BOMWorkload")

library(dplyr)
library(tidyr)
library(xlsx)
library(ggplot2)

# Make sure file is saved a comma separate value file.
df <- read.csv("E-CASCADE - Incoming Volumes with Revenue and Resubmit Reasons.csv",
               header=TRUE, sep=",")
employee <- read.xlsx("employee_region.xlsx", 1, header=TRUE)
headcount <- read.xlsx("production_hc.xlsx", 1, header=TRUE)
headCountReg <- filter(headcount, !Region=="Total") #Huh?????
headCountTotal <- filter(headcount, Region=="Total") #Huh????

# Get ride of unnecessary variables and rename columns.
dfReduced <- select(df, c(2, 13, 16,17,19,25,27,44,46))
names(dfReduced) <- c("ID", "Date", "Resubmit", "Audit", "Lead",
                      "Account", "TOV", "BOM", "Indicator")

dfReduced <- dfReduced %>%
    merge(employee, by.x="BOM", by.y="Name") %>%
    separate(col="BOM", into=c("BOM", "first1"), sep=",") %>%
    separate(col="Lead", into=c("Lead", "first2"), sep=",") %>%
    select(-c(first1, first2))

dfReduced$newDate <- dfReduced$Date %>%
    strptime(format=("%m/%d/%Y %H:%M")) %>%
    format("%y-%m") %>%
    as.factor()

rm(df)