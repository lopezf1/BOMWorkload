#setwd("~/RProgramming/BOMWorkload")

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

# Create BOM report for Resubmit Reason = Net New.
tempNetNew <- dfReduced %>% 
  filter(Resubmit=="1-Net New (Initial Request)") %>% 
  group_by(BOM, newDate) %>% 
  summarize(count=n()) %>% 
  spread(newDate, count) %>%
  as.data.frame()

rownames(tempNetNew) <- tempNetNew[,1] # Create row.names
tempNetNew[,1] <- NULL

tempNetNew$bomTotal <- rowSums(tempNetNew, na.rm=TRUE)
monthTotals <- colSums(tempNetNew, na.rm=TRUE)

reportNetNew <- rbind(tempNetNew, monthTotals)

len <- length(rownames(reportNetNew)) # Fix row name
rownames(reportNetNew)[len] <- "GrandTotal"

View(reportNetNew)

# Create BOM report for all Resubmit Reasons.
tempAll <- dfReduced %>% 
  group_by(BOM, newDate) %>% 
  summarize(count=n()) %>% 
  spread(newDate, count) %>%
  as.data.frame()

rownames(tempAll) <- tempAll[,1] # Create row.names
tempAll[,1] <- NULL

tempAll$bomTotal <- rowSums(tempAll, na.rm=TRUE)
monthTotals <- colSums(tempAll, na.rm=TRUE)

reportAll <- rbind(tempAll, monthTotals)

len <- length(rownames(reportAll)) # Fix row name
rownames(reportAll)[len] <- "GrandTotal"

View(reportAll)


# Create Regional report for Resubmit Reason = Net New
tempRegNetNew <- dfReduced %>% 
  filter(Resubmit=="1-Net New (Initial Request)") %>% 
  group_by(Region, newDate) %>% 
  summarize(count=n()) %>% 
  spread(newDate, count) %>%
  as.data.frame()

rownames(tempRegNetNew) <- tempRegNetNew[,1] # Create row.names
tempRegNetNew[,1] <- NULL

tempRegNetNew$regTotal <- rowSums(tempRegNetNew, na.rm=TRUE)
monthTotals <- colSums(tempRegNetNew, na.rm=TRUE)

reportRegNetNew <- rbind(tempRegNetNew, monthTotals)

len <- length(rownames(reportRegNetNew)) # Fix row name
rownames(reportRegNetNew)[len] <- "GrandTotal"

View(reportRegNetNew)


# Create Regional report for All Resubmit Reasons.

tempRegAll <- dfReduced %>% 
  group_by(Region, newDate) %>% 
  summarize(count=n()) %>% 
  spread(newDate, count) %>%
  as.data.frame()

rownames(tempRegAll) <- tempRegAll[,1] # Create row.names
tempRegAll[,1] <- NULL

tempRegAll$regTotal <- rowSums(tempRegAll, na.rm=TRUE)
monthTotals <- colSums(tempRegAll, na.rm=TRUE)

reportRegAll <- rbind(tempRegAll, monthTotals)

len <- length(rownames(reportRegAll)) # Fix row name
rownames(reportRegAll)[len] <- "GrandTotal"

View(reportRegAll)


# WORKLOAD PER MONTH

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

#PLOTTING: LOPEZ V MASSEY NET NEW

test <- dfReduced %>% 
  filter(Resubmit=="1-Net New (Initial Request)" &
           Lead=="MASSEY" | Lead=="LOPEZ") %>% 
  group_by(Lead, newDate) %>% 
  summarize(count=n())


g <- ggplot(test, aes(newDate, count))
g + geom_bar(stat="identity", aes(fill=Lead)) +
  geom_text(aes(label=count), size=3,
            hjust=0.5, vjust=3, position="stack") +
  ggtitle("New New Volumes by Lead") +
  ylab("Net New SR/WR Count") +
  xlab("YY-MM")

p <- ggplot(test, aes(newDate, count, color=Lead))
p + geom_line(aes(group=Lead)) +
    geom_point() +
    ggtitle("New New Volumes by Lead") +
    ylab("Net New SR/WR Count") +
    xlab("YY-MM")

# PLOTTING LOPEZ V MASSEY TOTAL

test2 <- dfReduced %>% 
    filter(Lead=="MASSEY" | Lead=="LOPEZ") %>% 
    group_by(Lead, newDate) %>% 
    summarize(count=n())

p <- ggplot(test2, aes(newDate, count, color=Lead))
p + geom_line(aes(group=Lead)) +
    geom_point() +
    ggtitle("Total Volumes by Lead") +
    ylab("Total SR/WR") +
    xlab("YY-MM")

# PLOTTING POLAK TOTAL VOLUMES
 
test3 <- dfReduced %>% 
    group_by(newDate) %>% 
    summarize(count=n())

p <- ggplot(test3, aes(newDate, count, color="red"))
p + geom_line(aes(group=1)) +
    geom_point() +
    ggtitle("Polak Total Volumes") +
    ylab("Total SR/WR") +
    xlab("YY-MM")
