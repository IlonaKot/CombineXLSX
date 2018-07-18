---
title: "Combining XLSX files into one!"
author: "Ilona Kotlewska"
date: "18 07 2018"
---
# This is a step-by-step tutorial on how to use this script.
# You might want to install RStudio to execute this code.
  
#SETUP
# First we need to install additional packages to R, which deal with Excel files
install.packages("XLConnect")

# Here I use the directory with all the data in one folder
setwd("~/Asya_2018/all_in_one")

# I create an empty object (in R an object is a data frame), because I will write the data into this empty object.
combined_file <- data.frame()

# This reads the file list from the folder
files = list.files(pattern="*.xlsx")

# Loops through all the files, reads them and concatenates
for (file in files) {
  single_file <- XLConnect::readWorksheetFromFile(file, sheet=2)
  combined_file <- rbind(combined_file, single_file)
}

# Saves a csv (stands for "comma separated file") to be opened in excel and changed into xlsx format.
write.csv(combined_file, file = 'data_combined.csv', row.names=F) 
