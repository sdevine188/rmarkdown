library(RDCOMClient)
library(fs)
library(rmarkdown)

# setwd
setwd("C:/Users/Stephen/Desktop/University of Wisconsin/classes/DS710/ds710spring2019assignment2")

# knit .rmd file
rmarkdown::render("assignment2_r.Rmd")


##############################################################


# convert word docx to pdf

# https://stackoverflow.com/questions/51957310/error-converting-docx-to-pdf-using-pandoc

# convert word docx to pdf, for example after knitting to docx
# IMPORTANT: it seems like the file path to the docx cannot have any spaces in it!!
# I got errors trying to convert
# "C:/Users/Stephen/Desktop/University of Wisconsin/classes/DS710/ds710spring2019assignment1/Assignment_1_R.docx"
# but no errors converting the same doc saved to
# "C:/Users/Stephen/Desktop/R/RDCOMClient/Assignment_1_R.docx"

# because file path can't have spaces, need to copy docx file to different folder before converting to pdf
path <- "assignment2_r.docx"
file_copy(path = path, new_path = "C:/Users/Stephen/Desktop/R/RDCOMClient/DS710/assignment2_r.docx")


#########################################################


# get new file path
file <- "C:/Users/Stephen/Desktop/R/RDCOMClient/DS710/assignment2_r.docx"

# use RDCOMClient to convert docx to pdf
wordApp <- COMCreate("Word.Application")  # create COM object
wordApp[["Visible"]] <- TRUE #opens a Word application instance visibly
wordApp[["Documents"]]$Add() #adds new blank docx in your application
wordApp[["Documents"]]$Open(Filename=file) #opens your docx in wordApp

#THIS IS THE MAGIC    
wordApp[["ActiveDocument"]]$SaveAs("C:/Users/Stephen/Desktop/R/RDCOMClient/DS710/assignment2_r.pdf", 
                                   FileFormat=17) #FileFormat=17 saves as .PDF

wordApp$Quit() #quit wordApp


#################################################


# confirm the pdf was created, and copy back to univ. of wisconsin folder
dir_ls("C:/Users/Stephen/Desktop/R/RDCOMClient/DS710")
file_copy(path = "C:/Users/Stephen/Desktop/R/RDCOMClient/DS710/assignment2_r.pdf",
          new_path = "assignment2_r.pdf")
dir_ls()





