library(dplyr)
library(readxl)
# library(googledrive)
# library(googlesheets4)
library(shiny)
library(shinydashboard)
library(shinydashboardPlus)
library(shinyRadioMatrix)
library(shinyWidgets)

rm(list = ls()) # clear environment

#################################### FUNCTIONS (set to directory of functionlibrary.R)
source("C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/scripts/functionlibrary.R", local = TRUE)

#################################### Set directories (set to directory of your template data)
#path = "C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/PRP-QUANT/PRP_QUANT_V3_21.xlsx"
#path = "C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/Deviation-Template/deviation_template_5.xlsx"
path = "C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/Scoping-Review/ScopingReview_4.xlsx"
sheets <- excel_sheets(path = path) #contains list of sheet names

# directory to save file
#setwd("C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/Deviation-Template")
setwd("C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/Scoping-Review")
#setwd("C:/Users/mueller_admin.ZPIDNB21/Documents/Desktop/Rprojects/PRP-QUANT/")

#set filename for file to save
#filename = "_prp11.RData"
filename = "_scr7.RData"
#filename = "_dev11.RData"

#################################### Google stuff
# Authentication
# Es macht mehr Sinn, einmal eine ganze Excel-Datei runterzuladen, als
# durch den spreadsheet auf dem Google Drive zu loopen (ist langsam).
 
# drive_auth(path = Sys.getenv("GOOGLE_SERVICE_ACCOUNT_KEY")) 
# gs4_auth(token = drive_token()) # synchronizes the tokens between sheets and drive
# url <- "https://docs.google.com/spreadsheets/d/1tsgNzlYbaw17i7B49Cz0Box0nWoNczRPw0v4cd7oqrA/edit#gid=587132800" # public template in stm account

# read spreadsheet
# googlesheet <- gs4_get(url) # fetches the metadata of that sheet with the url
# sheets <- googlesheet$sheets$name # get names of all sheets in the spreadsheet
# drive_download("PRP_QUANT_V2_itemtypes_sheet_12.xlsx")



#################################### create objects to fill (leave as is)
## dataframe that stores item IDs and item names 
ques_ans <- data.frame(ques = character(), ID = character(), cond = character(), meta = character(), stringsAsFactors = FALSE)

## list to store items of all sheets 
all_items <- list()

#################################### loop through sheets (leave as is)
for (m in seq_along(sheets)) {
  
  new_row <- data.frame(ques = paste("hxxd_", sheets[m], sep = ""), ID = "something", cond = "static", meta = "skip") # add row with page heading
 
  tabpanel <- read_excel(path, sheets[m])    # read single sheet (= all items that will later appear in one tabpanel of the app)
  # if you don't want to include the indeces ("T1 Title" instead of "Title"), make the next line a comment
  tabpanel$head <- with(tabpanel, paste(id, head)) # merge ID and heading, e.g., T1 Title
  
  mylist <- items_sheet(tabpanel) # items_sheet() is defined in functionlibrary.R
  all_items <- append(all_items, list(mylist))          # list that contains items of one section

  # exclude items of type "none" from report
  new_row_block <- subset(tabpanel, tabpanel$type != "none")  
  cond <- apply(new_row_block, 1, check_cond) # apply check_cond to every row of tabpanel

  new_row_block <- data.frame(ques = new_row_block$head, ID = new_row_block$id, cond = cond, meta = new_row_block$metadata)
  ques_ans <- bind_rows(ques_ans, new_row, new_row_block)
  
}

#################################### rename objects (set "template" to your preferred name of output files)
#template = "prp"
#template = "scope"
template = "temp"
#template = "dev" 
sheetsname <- paste(template, "_sheets", sep="") # e.g. "prp_sheets"
assign(sheetsname, sheets)

itemsname <- paste(template, "_items", sep="") # e.g. "prp_items"
assign(itemsname, all_items)

ques_ans_name <- paste(template, "_ques_ans", sep = "")
assign(ques_ans_name, ques_ans)

#################################### save objects (uncomment and set name filename first)
## save the relevant variables


variables_to_save <- c(sheetsname, itemsname, ques_ans_name)
save(list = variables_to_save, file = paste0(template, filename))

#################################### tidy up the workspace
rm(list = ls())
