require(dplyr)
require(readxl)
require(WriteXLS)

# setwd("../blog")

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

sheet_1 <- readSheetContentAsText(kobo.data.filename1, 1)
sheet_1 <- fixSheet(sheet_1)

sheet_2 <- readSheetContentAsText(kobo.data.filename1, 2)
sheet_2 <- fixSheet(sheet_2)

sheet_3 <- readSheetContentAsText(kobo.data.filename1, 3)
sheet_3 <- fixSheet(sheet_3)


WriteXLS(x=c("sheet_1","sheet_2","sheet_3"), kobo.data.outfilename)