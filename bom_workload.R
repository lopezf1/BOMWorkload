library(dplyr)
library(tidyr)
library(xlsx)
library(ggplot2)

# Make sure file is saved a comma separate value file.
df <- read.csv("Cascade (Polak 05192015 raw data) Incoming SRWR Volumes and Revenues with Resubmit Reasons.csv",
               header=TRUE, sep=",")
employee <- read.xlsx("employee_region.xlsx", 1, header=TRUE)
headcount <- read.xlsx("production_hc.xlsx", 1, header=TRUE)
headCountReg <- filter(headcount, !Region=="Total")
headCountTotal <- filter(headcount, Region=="Total")

# Get ride of unnecessary variables and rename columns.
dfReduced <- select(df, c(1,2,16,17,19,25,27,44,46))
names(dfReduced) <- c("Date", "ID", "Resubmit", "Audit", "Lead",
                "Account", "TOV", "Owner", "Indicator")
dfReduced <- merge(dfReduced, employee, by.x="Owner", by.y="Name")
rm(df)

# Create BOM report for Resubmit Reason = Net New.
tempNetNew <- dfReduced %>% 
  filter(Resubmit=="1-Net New (Initial Request)") %>% 
  group_by(Owner, Date) %>% 
  summarize(count=n()) %>% 
  spread(Date, count) %>%
  as.data.frame()

rownames(tempNetNew) <- tempNetNew[,1] # Create row.names
tempNetNew[,1] <- NULL

tempNetNew$bomTotal <- rowSums(tempNetNew, na.rm=TRUE)
monthTotals <- colSums(tempNetNew, na.rm=TRUE)

reportNetNew <- rbind(tempNetNew, monthTotals)

len <- length(rownames(reportNetNew)) # Fix row name
rownames(reportNetNew)[len] <- "GrandTotal"

View(reportNetNew)

# # Create BOM report for Resubmit Reason = All
tempAll <- dfReduced %>% 
  group_by(Owner, Date) %>%
  summarize(count=n()) %>%
  spread(Date, count) %>%
  as.data.frame()

rownames(tempAll) <- tempAll[,1] # Create row.names
tempAll[,1] <- NULL

tempAll$bomTotal <- rowSums(tempAll, na.rm=TRUE)
monthTotals <- colSums(tempAll, na.rm=TRUE)

reportAll <- rbind(tempAll, monthTotals)

len <- length(rownames(reportAll)) # Fix row name
rownames(reportAll)[len] <- "GrandTotal"

View(reportAll)

# Write report results to single leader file.
write.xlsx(reportNetNew, file="polak_workload.xlsx", sheetName="NetNew")
write.xlsx(reportAll, file="polak_workload.xlsx", sheetName="All", append=TRUE)

# Create Regional report for Resubmit Reason = Net New
tempRegNetNew <- dfReduced %>% 
  filter(Resubmit=="1-Net New (Initial Request)") %>% 
  group_by(Region, Date) %>% 
  summarize(count=n()) %>% 
  spread(Date, count) %>%
  as.data.frame()

rownames(tempRegNetNew) <- tempRegNetNew[,1] # Create row.names
tempRegNetNew[,1] <- NULL

tempRegNetNew$bomTotal <- rowSums(tempRegNetNew, na.rm=TRUE)
monthTotals <- colSums(tempRegNetNew, na.rm=TRUE)

reportRegNetNew <- rbind(tempRegNetNew, monthTotals)

len <- length(rownames(reportRegNetNew)) # Fix row name
rownames(reportRegNetNew)[len] <- "GrandTotal"

View(reportRegNetNew)

# Create Regional report for Resubmit Reason = All
tempRegAll <- dfReduced %>% 
  group_by(Region, Date) %>% 
  summarize(count=n()) %>% 
  spread(Date, count) %>%
  as.data.frame()

rownames(tempRegAll) <- tempRegAll[,1] # Create row.names
tempRegAll[,1] <- NULL

tempRegAll$bomTotal <- rowSums(tempRegAll, na.rm=TRUE)
monthTotals <- colSums(tempRegAll, na.rm=TRUE)

reportRegAll <- rbind(tempRegAll, monthTotals)

len <- length(rownames(reportRegAll)) # Fix row name
rownames(reportRegAll)[len] <- "GrandTotal"

View(reportRegAll)

# Write report results to single leader file.
write.xlsx(reportRegNetNew, file="polak_workload.xlsx",
           sheetName="RegNetNew", append=TRUE)
write.xlsx(reportRegAll, file="polak_workload.xlsx",
           sheetName="RegAll", append=TRUE)

# Determine workload per month

tempPerReg <- dfReduced %>% 
    group_by(Region, Date) %>% 
    summarize(count=n()) %>% 
    as.data.frame() %>%
    cbind(hc=headCountReg$HeadCount) %>%
    mutate(load=count/hc) %>%
    select(-(count:hc)) %>%
    spread(Date, load)

rownames(tempPerReg) <- tempPerReg[,1] # Create row.names
tempPerReg[,1] <- NULL

#TESTING STARTING HERE
tempPerAll <- dfReduced %>%
    group_by(Region, Date) %>% 
    summarize(count=n()) %>% 
    as.data.frame() %>%
    spread(Date, count)

rownames(tempPerAll) <- tempPerAll[,1] # Create row.names
tempPerAll[,1] <- NULL

monthTotals <- colSums(tempPerAll, na.rm=TRUE) %>%
  cbind(hc=headCountTotal$HeadCount)
colnames(monthTotals) <- c("total", "hc")
monthTotals <- as.data.frame(monthTotals)
monthTotals <- mutate(monthTotals, load=total/hc)

View(monthTotals)





