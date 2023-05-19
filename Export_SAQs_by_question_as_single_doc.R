library(officer) # for making Word docs 
library(readxl) # for reading Excel file
library(svDialogs)
library(tidyverse) # for tidy code 
library(janitor) # for cleaning variable names 
library(tools)


# Use interactive dialogue to get the name of the results file 
Responses <- dlg_open(
  title = "Please pick the file containing the quiz responses",
  multiple = FALSE,
  filters = dlg_filters[c("xls", "All"), ],
  gui = .GUI
)

# Get the base file name (not including directory) and strip off the file extension 
fName <- basename(Responses$res)
saveName <- file_path_sans_ext(fName)

# Import responses from Excel 
All_responses <- read_excel(fName)
# Reads the file that has your responses (make sure you have the correct name here)
All_responses <- All_responses %>%
  clean_names() # Remove spaces in names
rows <- nrow(All_responses) # Get number of rows

SAQ1 = readline(prompt = "What is the number of your first SAQ? ")
SAQ2 = readline(prompt = "What is the number of your last SAQ? ")

SAQ1 <- as.integer(SAQ1)
SAQ2 <- as.integer(SAQ2)

questions <- SAQ1:SAQ2


nQuestions <- length(questions) # Work out how many this is 
varlistQ  <- paste("Question", questions[1]:questions[nQuestions], sep = "_") 
# Make a list of the question names (for the doc)
varlistR <- paste("response", questions[1]:questions[nQuestions], sep = "_")
# Make a list of the response names (to extract from the columns)

docName <- paste(saveName," SAQs.docx", sep = "") # Document name (no spaces)
doc1 <- read_docx()  # calls function to create a document using officeR
# doc1 <- body_add_title(doc1, saveName, 1)
# Loop that makes a separate document for each question
for (i in 1:nQuestions){
  qName <- varlistQ[i] # Question name
  doc1 <- body_add_par(doc1, value = qName, style = "heading 1") # Add question name at top
  for (q in 1:rows) # student loop (could also make this anonymous by using student no.)
  {Name <- paste(All_responses$first_name[q], All_responses$surname[q], sep = " ") # Student name
    doc1 <- body_add_par(doc1, value = Name , style = "heading 2", pos = "after") # Add student name 
    doc1 <- body_add_par(doc1, value = All_responses[[varlistR[i]]][q], style = "Normal", pos = "after") # add their response to this question
    } 
}

print(doc1, target = docName) # Save document 
    
  