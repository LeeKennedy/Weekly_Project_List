# Clean Up environment ---------------------------------------------------
rm(list=ls())

# Packages ---------------------------------------------------------------
library(readxl)
library(readr)
library(dplyr)
library(tidyr)
library(ggplot2)
library(xlsx)

# Data Input -------------------------------------------------------------


if("Windows" %in% Sys.info()['sysname'] == TRUE){ 
        setwd() 
} else { 
        setwd("~/Documents/GitHub/Weekly_Project_List/output")
}

file.remove("Weekly_Old.xlsx")
file.copy("Weekly.xlsx", "Weekly_Old.xlsx")
file.remove("Weekly.xlsx")

if("Windows" %in% Sys.info()['sysname'] == TRUE){ 
        new_list <- loadWorkbook("~/Documents/GitHub/Weekly_Project_List/data/Weekly_List_Template.xlsx")
        new_input <- read.csv("~/Desktop/Project List/project_list.csv")
        old_list <- read_excel("~/Desktop/Project List/Weekly_Old.xlsx") 
} else { 
        new_list <- loadWorkbook("~/Documents/GitHub/Weekly_Project_List/data/Weekly_List_Template.xlsx")
        new_input <- read.csv("~/Documents/GitHub/Weekly_Project_List/output/project_list.csv")
        old_list <- read_excel("~/Documents/GitHub/Weekly_Project_List/output/Weekly_Old.xlsx")
}




old_list <- old_list[,c(3,4)]

new_1 <- merge(x=new_input, y= old_list, by.x = "Project", by.y = "Project", all = TRUE)

new_1 <- new_1[,c(2,3,1, 11, 5:10)]

colnames(new_1)[4] <- "Comment"

new_1[is.na(new_1)] <- ""

new_1 <- new_1 %>% 
        arrange(key)
new_1 <- new_1[-1,]

new_s <- getSheets(new_list)

addDataFrame(new_1, sheet = new_s$List, row.names = FALSE, col.names = FALSE, startRow=2, startColumn = 1) 


# write data to sheet starting on line 1, column 4
saveWorkbook(new_list, "Weekly.xlsx")

temp_list <- read_excel("~/Documents/GitHub/Weekly_Project_List/output/Weekly.xlsx")

temp_list <- temp_list[,c(2:4)]
temp_list[is.na(temp_list)] <- ""

temp_list <- temp_list %>% 
        filter(Area != "Done")

write.xlsx(temp_list, "LK_weekly_list.xlsx")





