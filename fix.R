require(dplyr)
require(readxl)
require(WriteXLS)
getwd()
setwd("./blog")

choices <- read_excel("data/questions.xls",sheet="choices")
choices[,c("a6","a75")] <- list(NULL)

kobo.data.outfilename = "final.xlsx"
kobo.data.filename1 = "data/file1.xlsx"
kobo.data.filename2 = "data/file2.xlsx"

readSheetContentAsText = function (kobo.data.filename, sheet_number) {
  # read sheet, which gives warnings like `expecting numeric: got 'a46-1'` and these data were missing from df
  sheet_mdd <- read_excel(kobo.data.filename, sheet=sheet_number)
  # find the length of columns so that we can create types and re-read the sheet with text-type for all the columns
  #   to avoid these warnings
  no.of.column = length(colnames(sheet_mdd))
  col_types = rep("text", no.of.column)
  # https://stackoverflow.com/questions/33353563/read-excel-expecting-numeric-and-value-is-numeric
  
  return(read_excel(kobo.data.filename, sheet=sheet_number, col_types = col_types))
}

fixSheet = function(sheet) {
  for(colname in colnames(sheet)) {
    # for each column, find the first non-empty value
    firstNonEmptyRowVal = head(na.omit(sheet[colname]),1)[[colname]]
    # see if the value is present in the choices$name; if yes apply the mapping
    if(!identical(firstNonEmptyRowVal, character(0)))
      if(firstNonEmptyRowVal %in% choices$name) {
        sheet[colname] = (choices[match(sheet[[colname]], choices$name),3])
      }
  }
  sheet <- data.frame(lapply(sheet, trimws))
  # https://stackoverflow.com/questions/20760547/removing-whitespace-from-a-whole-data-frame-in-r
  sheet <- data.frame(lapply(sheet, function(x)gsub('\n',' ',x)))
  return(sheet);
}

# read and fix first file 
sheet1.1 <- readSheetContentAsText(kobo.data.filename1, 1)
sheet1.1 <- fixSheet(sheet1.1)

sheet1.2 <- readSheetContentAsText(kobo.data.filename1, 2)
sheet1.2 <- fixSheet(sheet1.2)

sheet1.3 <- readSheetContentAsText(kobo.data.filename1, 3)
sheet1.3 <- fixSheet(sheet1.3)

# read and fix second file
sheet2.1 <- readSheetContentAsText(kobo.data.filename2, 1)
sheet2.1 <- fixSheet(sheet2.1)

sheet2.2 <- readSheetContentAsText(kobo.data.filename2, 2)
sheet2.2 <- fixSheet(sheet2.2)

sheet2.3 <- readSheetContentAsText(kobo.data.filename2, 3)
sheet2.3 <- fixSheet(sheet2.3)

# merge two sheets data from 2 files into one dataframe
sheet1.1 <- rbind(sheet1.1,sheet2.1)
sheet1.2 <- rbind(sheet1.2,sheet2.2)
sheet1.3 <- rbind(sheet1.3,sheet2.3)

# write to XLS
WriteXLS(x=c("sheet1.1","sheet1.2","sheet1.3"), kobo.data.outfilename)