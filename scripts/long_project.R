# Clean Up environment ---------------------------------------------------
rm(list=ls())

# Packages ---------------------------------------------------------------
library(readxl)
library(readr)
library(dplyr)
library(tidyr)
library(stringr)

# Data Input -------------------------------------------------------------

lk_proj <- read_csv("~/Desktop/Project List/LK_Projects_full.csv")
#lk_proj <- read_csv("C:/Users/leekennedy/Desktop/Project List May 2017/LK_Projects_full.csv")


# Data Cleaning ----------------------------------------------------------

tidy_proj <- tidyr::gather(data = lk_proj, key = Area, value = Project, na.rm = FALSE, `To Do`, Doing, `With Lab`, `With IT`, `With Quality`, Limbo, Done)

tidy_proj <- na.omit(tidy_proj)

### Add sorting key ------------------------------------------------------

tidy_proj$key <- 0
tidy_proj <- tidy_proj[,c(3,1,2)]
n <- nrow(tidy_proj)

for (i in 1:n) {
        if(tidy_proj$Area[i] == "Doing") tidy_proj$key[i] = 1
        if(tidy_proj$Area[i] == "With Lab") tidy_proj$key[i] = 2
        if(tidy_proj$Area[i] == "With Quality") tidy_proj$key[i] = 3
        if(tidy_proj$Area[i] == "With IT") tidy_proj$key[i] = 4
        if(tidy_proj$Area[i] == "Limbo") tidy_proj$key[i] = 5
        if(tidy_proj$Area[i] == "To Do") tidy_proj$key[i] = 6
        if(tidy_proj$Area[i] == "Done") tidy_proj$key[i] = 7
}




### Sort and initial split ---------------------------------------------

tidy_proj <- tidy_proj %>% 
        arrange(key, Project) %>%
        separate(Project, c("Project","Control"),"\n") 

### Split of the control markers ---------------------------------------

tidy_proj$GB <- NA
tidy_proj$Project_No <- NA
tidy_proj$CC <- NA
tidy_proj$NM <- NA
tidy_proj$EM <- NA
tidy_proj$Story <- NA

for (i in 1:n){
        
        temp <- unlist(strsplit(tidy_proj$Control[i], ", "))
        m = length(temp)
        
        if (m == 0) next
        
        for (j in 1:m) {
        if(grepl("GB", temp[j]) == TRUE) tidy_proj$GB[i] = "GB"
        if(grepl("Project", temp[j]) == TRUE) tidy_proj$Project_No[i] = str_sub(temp[j], start= -3)
        if(grepl("CC", temp[j]) == TRUE) tidy_proj$CC[i] = str_sub(temp[j], start= -3)
        if(grepl("NM", temp[j]) == TRUE) tidy_proj$NM[i] = str_sub(temp[j], start= -5)
        if(grepl("EM", temp[j]) == TRUE) tidy_proj$EM[i] = str_sub(temp[j], start= -5)
        if(grepl("Story", temp[j]) == TRUE) tidy_proj$Story[i] = str_sub(temp[j], start= -3)
        }
}

tidy_proj[is.na(tidy_proj)] <- ""
colnames(tidy_proj)[4] <- "Comment"
tidy_proj$Comment <- ""

# Exporting Data -------------------------------------------------------

write_csv(tidy_proj, "project_list.csv")





