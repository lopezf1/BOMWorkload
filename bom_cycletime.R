# Just focus on work cycle-time.  Assume all inbox cycle-time
# handled wihtin one day.

setwd("~/RProgramming/BOMWorkload")

library(dplyr)
library(tidyr)

# Read in employee file for regional analysis.

employee <- read.csv("employee_region.csv",
                     header=TRUE, sep=",") %>%
    separate(col="Name", into=c("BOM", "first"), sep=",") %>%
    select(-first)

# Read in Work Cycle-time file, select specific columns and rename.

dfWork <- read.csv("E-CASCADE - Work Cycle Time.csv",
               header=TRUE, sep=",") %>%
    select(c(7, 4, 13, 19, 22, 5))
names(dfWork) <- c("ID", "Date", "Status", "Lead",
                   "BOM", "Interval")

# Simplify BOM names and combine employee and cycle-time files.

dfWork <- dfWork %>%
separate(col="BOM", into=c("BOM", "first1"), sep=",") %>%
separate(col="Lead", into=c("Lead", "first2"), sep=",") %>%
select(-c(first1, first2)) %>%
merge(employee, by.x="BOM", by.y="BOM")  

# Fix up the dates.

dfWork$newDate <- dfWork$Date %>%
strptime(format=("%m/%d/%Y %H:%M")) %>%
format("%y-%m") %>%
as.factor()

# Now create a table of values.

# Create plot for workcycle-time in July.
box <- dfWork %>% 
    filter(newDate=="15-07" & Region=="Named") %>% 
    as.data.frame()

boxplot(Interval ~ Region, data=box)
str(box)
