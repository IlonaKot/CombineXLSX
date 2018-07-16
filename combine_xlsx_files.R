---
title: "Combining XLSX files into one!"
author: "Ilona Kotlewska"
date: "16 07 2018"
---
# This is a step-by-step tutorial on how to use this script.
# You might want to install RStudio to execute this code.
  
#SETUP
# First we need to install additional packages to R, which deal with Excel files
install.packages("XLConnect")

# A working directory is a folder in which the script and the result file will be saved, and where the data are stored.
# We set the working directory with "setwd" command. YOUR DIRECTORY MIGHT BE DIFFERENT! PUT HERE THE PATH TO YOUR DATA :)
setwd("~/Asya_2018/")

# I write the folder location in a variable called "file_dir", because I will use it later.
# (PUT HERE THE SAME AS YOU DID ABOVE)
file_dir <- "~/Asya_2018/"

# I create an empty object (in R an object is a data frame), because I will write the data into this empty object.
combined_file <- data.frame()

# Now I declare the list of all the folders with the data
folder_list <- c("04-12-2018/", "04-28-2018/", "05-14-2018/", "05-30-2018/", "06-15-2018/", "07-01-2018/",
                   "04-13-2018/", "04-29-2018/", "05-15-2018/", "05-31-2018/", "06-16-2018/", "07-02-2018/",
                   "04-14-2018/", "04-30-2018/", "05-16-2018/", "06-01-2018/", "06-17-2018/", "07-03-2018/",
                   "04-15-2018/", "05-01-2018/", "05-17-2018/", "06-02-2018/", "06-18-2018/", "07-04-2018/",
                   "04-16-2018/", "05-02-2018/", "05-18-2018/", "06-03-2018/", "06-19-2018/", "07-05-2018/",
                   "04-17-2018/", "05-03-2018/", "05-19-2018/", "06-04-2018/", "06-20-2018/", "07-06-2018/",
                   "04-18-2018/", "05-04-2018/", "05-20-2018/", "06-05-2018/", "06-21-2018/", "07-07-2018/",
                   "04-19-2018/", "05-05-2018/", "05-21-2018/", "06-06-2018/", "06-22-2018/", "07-08-2018/",
                   "04-20-2018/", "05-06-2018/", "05-22-2018/", "06-07-2018/", "06-23-2018/", "07-09-2018/",
                   "04-21-2018/", "05-07-2018/", "05-23-2018/", "06-08-2018/", "06-24-2018/", "07-10-2018/",
                   "04-22-2018/", "05-08-2018/", "05-24-2018/", "06-09-2018/", "06-25-2018/", "07-11-2018/",
                   "04-23-2018/", "05-09-2018/", "05-25-2018/", "06-10-2018/", "06-26-2018/", "07-12-2018/",
                   "04-24-2018/", "05-10-2018/", "05-26-2018/", "06-11-2018/", "06-27-2018/",
                   "04-25-2018/", "05-11-2018/", "05-27-2018/", "06-12-2018/", "06-28-2018/",
                   "04-26-2018/", "05-12-2018/", "05-28-2018/", "06-13-2018/", "06-29-2018/",
                   "04-27-2018/", "05-13-2018/", "05-29-2018/", "06-14-2018/", "06-30-2018/")

# EXECUTE!
# This is an all in one nested loop that does everything described in the STEP-BY-STEP part

for (folder in folder_list) {
  file_path <- paste(file_dir, folder, sep='')
  for (file in dir(file_path, full.names=T)) {
    single_file <- XLConnect::readWorksheetFromFile(file, sheet=2)
  }
  combined_file <- rbind(combined_file, single_file)
}



# LEARNING TIME!

# STEP BY STEP
# Now step by step to explain what has been done and how it happend. Take a look at single stages.
# Do not execute these commands if you want to run the script again. Go to the "EXECUTE" part instead!
# Otherwise the step by step will process everything on only one file instead of all, because the loops here were separated from
# each other and do not store the information about all the files.

# 1) We need to read file, right? So lets go to a one particular folder, let's say "04-12-2018/":
setwd("~/Asya_2018/04-12-2018/")
# And let's read the file in it. 
# We need to use the library name first, XLConnect::here_add_the_name_of_the_function("Here_the_filename", sheet_nr)
# In your Excel files the first sheet contains figures and the second sheet contains data, that's why we use the argument sheet = 2)
XLConnect::readWorksheetFromFile("AGL_Refill_04-12-2018.xlsx", sheet=2)
# This should give an output like this:
#Date A1.Volume A1.Valve.Status A2.Volume A2.Valve.Status A3.Volume A3.Valve.Satus A4.Volume A4.Valve.Status
#1    2018-04-12 04:00:00         0               0        72               0        82              0        60               0
#2    2018-04-12 04:01:00         0               0        72               0        82              0        60               0
#... and so on.

# Ok, now as we know that our command to read the single file works, we save the file to an object, let's call it "single file":
single_file <- XLConnect::readWorksheetFromFile("AGL_Refill_04-12-2018.xlsx", sheet=2)

# 2) Looping througn the folders
# We need to make sure our script will look into all the folders, that's why we need to create a loop, which will go through all the folder names.
# This goes to all the folder names defined in the "folder_list" and writes their names together with the path to a variable called "file_path".
for (folder in folder_list) {
  file_path <- paste(file_dir, folder, sep='')
}
# To make sure this loop works properly you may take a look at the variable file_path. It should look more less like this:
# "~/Asya_2018/06-30-2018"
# Why more less? Because your folder where you store the data might be different (and you should define it in the beginning of this
# script. And, because my folder list ends with the subject 06-30-2018, but yours might be organized differently, for example the
# last subject might be "07-12-2018/").

# 3) Reading file in a loop
# This loop goes file by file created in the file_path variable and reads single files.
for (file in dir(file_path, full.names=T)) {
  single_file <- XLConnect::readWorksheetFromFile(file, sheet=2)
}

# 4) Combine files
# Finally we need to combine all the read files by adding every one of them to the empty data frame created above.
combined_file <- rbind(combined_file, single_file)

# Good luck! :)