
require("rmarkdown")
require("knitr")
require("grid")
#require("xlsx")
require("tframe")
require("tframePlus")
require("lubridate")
require("stringr")
require("scales")
require("zoo")
require("xts")

require("forecast")
require("car")
require("reshape2")
require("ggplot2")
require("tidyr")
require("plyr") #Hadley said if you load plyr first it should be fine
require("dplyr")
library("readr")
require("XLConnect")


#######
#
# reads in the "Pipeline by Year Open" tab from multiple pipeline summary files
# saves an Rdata file as a result

# test XLconnect by creating a new workbook
wb.new <- loadWorkbook("myNewExcelFile.xlsx", create = TRUE)


#######
#
# name of files to read


yr_list <- rep(2007:2012)
yr_list <- c(paste("input_data/str_files/Pipeline_December ", yr_list, ".xls", sep=""))
# start in 2008 for the May files because the tab names are different in the May 2007 version
yr_list2 <- rep(2008:2013)
yr_list2 <- c(paste("input_data/str_files/Pipeline_May ", yr_list2, ".xls", sep=""))
# read in May 2007 separately given different format
yr_list3 <- rep(2007:2007)
yr_list3 <- c(paste("input_data/str_files/Pipeline_May ", yr_list3, " extract.xls", sep=""))

newyr_list <- rep(2013:2019)
newyr_list <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list, "12.xls", sep=""))
newyr_list2 <- rep(2014:2019)
newyr_list2 <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list2, "05.xls", sep=""))
newyr_list3 <- rep(2014:2019)
newyr_list3 <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list3, "09.xls", sep=""))
newyr_list4 <- rep(2013:2019)
newyr_list4 <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list4, "11.xls", sep=""))
newyr_list5 <- rep(2015:2019)
newyr_list5 <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list5, "03.xls", sep=""))
newyr_list6 <- rep(2018:2020)
newyr_list6 <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list6, "04.xls", sep=""))
newyr_list7 <- rep(2015:2018)
newyr_list7 <- c(paste("input_data/str_files/PipelineSummary_US_", newyr_list7, "06.xls", sep=""))

yr_list <- c(yr_list, yr_list2, yr_list3, newyr_list, newyr_list2, newyr_list3, newyr_list4, newyr_list5, newyr_list6)
yr_list

# Create data frame with NA's
out_open_year <- data.frame(matrix(NA, nrow=0, ncol=13)) 
colnames(out_open_year) <- c("sourcemonth", 
             "open_year", 
             "segment",
             "rmsincns",
             "rmsfnlpl",  
             "rmsplnnn",
             "rmstotal",
             "rmsprpln",
             "htlsincns",
             "htlsfnlpl",
             "htlsplnnn",
             "htlstotal",
             "htlsprpln")

for (y in yr_list) {
#y <- c("input_data/Pipeline_May 2007 extract.xls")
#y <- c("input_data/str_files/PipelineSummary_US_201903.xls")
    print(paste("starting ", y, sep=""))
  # I started with read xlsx package but I wasn't reading the formulas used for the totals
  # starting in 2013 files. So I switched over to XLConnect, which happened to work
  #tempa <- read.xlsx(y, sheetName="Pipeline by Year Open", startRow=1,colIndex =1:14,
  #                   header = FALSE)
  wb = loadWorkbook(y)
  tempa <- readWorksheet(wb, sheet="Pipeline by Year Open", startRow=1, endRow=57, startCol=1, 
                         endCol=14,
                         header = FALSE)
  temp <- tempa
  
    # removes rows that are all NA
    temp <-  temp[!!rowSums(!is.na(temp)),]
    
    # adds a new column at the beginning
    temp <- temp %>%
      mutate(open_year="NA") %>%
      rename(segment=Col1)
    
    ###########
    #
    # uses first column to put open date year into open_year current row and the next 11
    
    # creates list of years
    years_list <- rep(2007:2023)
    years_list <- paste("Open Date ", years_list, sep="")
    
    for (y in years_list) {
      # looks in second column and finds row with Open Date xxxx, for example
      a <- grep(y, as.character(temp$segment))
      # determines rows that will receive insertion
      op_yr_label <- rep(0:10)
      op_yr_label <- op_yr_label + a
      # grabs the text that will be inserted (should be Open Date xxx)
      op_yr <- as.character(temp[a,"segment"])
      # does the insertion
      for (i in op_yr_label) {
        temp[i,"open_year"] <-op_yr
      }
    }
    
    ############
    #
    # uses column 13 to contain the date of the pipeline summary file being read
    
    # finds row with text string
    a <- grep("Data for end of", as.character(temp$Col12))
    # puts contents of text string in holder
    src <- as.character(temp[a,"Col12"])
    # sometimes there is an asterix  
    src <- sub("\\*", "", src) 
    
    # puts src into the sourcemonth
    temp <- temp %>%
      mutate(sourcemonth=src)
    
    #############
    #
    # does some set up to get to column names
    
    # finds first row with Chain Scale as text in column 2
    a <- grep("Chain Scale", as.character(temp$segment))
    a <- a[1]
    
    # do the number of hotel projects first
    b <- as.matrix(temp[a,2:6]) %>%
      sub("-", "", .) %>% # replace space with '_'
      tolower() %>% # put in lower case
      abbreviate(., minlength = 5) %>% # creates short names
      sub("prpln", "uncnf", .) %>% # replace preplanning with unconfirmed
      paste("htls", ., sep="")
    
    # now use these column names
    d <- colnames(temp) %>%
      as.matrix()
    d[2:6] <- as.character(b) # puts text vector into d
    colnames(temp) <- d # uses d as column names
    
    
    # do the number of rooms
    c <- as.matrix(temp[a,8:12]) %>%
      sub("-", "", .) %>%
      tolower() %>%
      abbreviate(., minlength = 5) %>%
      sub("prpln", "uncnf", .) %>% # replace preplanning with unconfirmed  
      paste("rms", ., sep="")
    # now use these column names
    d <- colnames(temp) %>%
      as.matrix()
    d[8:12] <- as.character(c)
    colnames(temp) <- d
    
    #################
    #
    # drop a few columns and rows
    
    tempb <- temp
    
    tempb$Col7 <- NULL
    
    # get ready to filter
    target <- c("Luxury", "Upper Upscale", "Upscale", "Upper Midscale", "Midscale", 
                "Midscale w/ F&B", "Midscale w/o F&B", "Economy",
                "Unaffiliated", "Total")
    tempb <- tempb %>%
      filter(segment %in% target) %>% 
      mutate(segment = gsub("Total", "totus", segment)) %>%
      mutate(segment = gsub("Luxury", "luxus", segment)) %>%
      mutate(segment = gsub("Upper Upscale", "upuus", segment)) %>%
      mutate(segment = gsub("Upscale", "upsus", segment)) %>%
      mutate(segment = gsub("Upper Midscale", "upmus", segment)) %>%
      mutate(segment = gsub("Midscale w/ F&B", "midwus", segment)) %>%
      mutate(segment = gsub("Midscale w/o F&B", "midwous", segment)) %>%
      mutate(segment = gsub("Midscale", "midus", segment)) %>%
      mutate(segment = gsub("Economy", "ecous", segment)) %>%
      mutate(segment = gsub("Independents", "indus", segment)) %>%
      mutate(segment = gsub("Unaffiliated", "indus", segment)) %>%
      select(sourcemonth, open_year, segment, starts_with('rms'), starts_with('htls')) 
    
    # removes commas in columns 4 to 13
    tempb[, 4:13] <- apply(tempb[, 4:13], 2, function(x) as.numeric(gsub(",", "", x)))
    
    # checks classes of columns then converts certin columns to numeric and character
    sapply(tempb, class)
    # needed to convert from factors to character, and then to numeric
    # otherwise trying to go from factor directly to numeric had zeros coding as ones
    tempb[4:13] <- sapply(tempb[4:13],as.character)
    tempb[4:13] <- sapply(tempb[4:13],as.numeric)
    tempb[1:3] <- sapply(tempb[1:3],as.character)
    sapply(tempb, class)
    
    # replace NAs in certain columns with 0
    tempb[, 4:13][is.na(tempb[, 4:13])] <- 0
  
  out_open_year <- rbind(out_open_year, tempb)
}

backup <- out_open_year

out_open_year <- backup
out_open_year <- out_open_year %>%
  mutate(sourcemonth=sub("Data for end of ", "", sourcemonth)) %>%
  mutate(sourcemonth=sub(",", "", sourcemonth)) %>%
  mutate(sourcemonth=as.Date(as.yearmon(sourcemonth, "%B %Y"))) %>%
  mutate(srcm_month = month(sourcemonth, label=TRUE, abbr=TRUE)) %>%
  mutate(and_later=str_detect(open_year, "and Later")) %>%
  mutate(open_year=sub("Open Date ", "", open_year)) %>%
  mutate(open_year=sub(" and Later", "", open_year)) %>%
  mutate(open_year=paste(open_year,"-01-01", sep="")) %>%
  mutate(open_year=as.Date(open_year, "%Y-%m-%d")) %>%
  mutate(hori=(year(open_year) - year(sourcemonth))) %>%
  mutate(horitext=hori) %>%
  mutate(horitext=sub("0", "Current year", horitext)) %>%
  mutate(horitext=sub("1", "Yr1 (Next year)", horitext)) %>%
  mutate(horitext=sub("2", "Yr2 (Year after next)", horitext)) %>%
  mutate(horitext=sub("3", "Yr3", horitext)) %>%
  mutate(horitext=sub("4", "Yr4", horitext))
  
out_open_year$and_later <- Recode(out_open_year$and_later, "c(TRUE)=' and later'; c(FALSE)=''")
out_open_year <- out_open_year %>%
  mutate(horitext=paste(horitext, and_later, sep="")) %>%
  select(-and_later) %>%
  select(sourcemonth, srcm_month, open_year, hori, horitext, segment:htlsuncnf)

# saves Rdata version of the data
save(out_open_year, file="output_data/out_open_year.Rdata")

