
#######
#
# reads in the "Pipeline by Year Open" tab from multiple pipeline summary files
# saves an Rdata file as a result

#######
#
# name of files to read

yr_list <- rep(2007:2012)
yr_list <- c(paste("input_data/Pipeline_December ", yr_list, ".xls", sep=""))
newyr_list <- rep(2013:2015)
newyr_list <- c(paste("input_data/PipelineSummary_US_", newyr_list, "12.xls", sep=""))
yr_list <- c(yr_list, newyr_list)
yr_list

# Create Data Frame with NA's
out_changes <- data.frame(matrix(NA, nrow=0, ncol=19)) 

for (y in yr_list) {
  
  wb = loadWorkbook(y)
  tempa <- readWorksheet(wb, sheet="Existing Supply Chgs by Brand", startRow=1, startCol=1, 
                         endCol=21,
                         header = FALSE)
  temp <- tempa
  
  # looks in the first column to see if there any matches for the text string
  # if there are, set that element to NA. 
  e <- grep("SMITH TRAVEL RESEARCH", temp[,1], value=FALSE)
  for (i in e) {
    temp[i,1] <- NA
  }
  # same as above for another text string
  e <- grep("Smith Travel Research", temp[,1], value=FALSE)
  for (i in e) {
    temp[i,1] <- NA
  }
  # same as above for another text string
  e <- grep("STR, Inc.", temp[,1], value=FALSE)
  for (i in e) {
    temp[i,1] <- NA
  }
  # same as above for another text string
  e <- grep("tracking variances", temp[,1], value=FALSE)
  for (i in e) {
    temp[i,1] <- NA
  }
  # same as above for another text string
  e <- grep("corporate feed activity", temp[,1], value=FALSE)
  for (i in e) {
    temp[i,1] <- NA
  }
  
  # removes any columns that are all NA
  # this helps get rid of the first column in some source files
  # that come in because the source line is in the column
  # in other months the read in process just skips that first column
  temp <- temp[,colSums(is.na(temp))<nrow(temp)]
  
  # renames the columns so that they are easier to refer to
  x <- rep(1:ncol(temp))
  x <- paste("X", x, sep="")
  colnames(temp) <- x
  
  # removes rows that are all NA
  temp <-  temp[!!rowSums(!is.na(temp)),]
  
  # grabs the month of the source file from the header
  e <- grep("Existing", temp[,10], value=FALSE)
  
  src <- (temp[e-1,10]) %>%
    as.character %>%
    as.Date %>%
    as.yearmon %>%
    as.Date
  
  # createes a new column with the source file date
  temp <- temp %>%
    mutate(sourcemonth=src)
  
  #############
  #
  # does some set up to get to column names
  
  e <- grep("Existing", temp[,2], value=FALSE)
  f <- as.matrix(temp[e-2,1:18])
  g <- as.matrix(temp[e,1:18])
  h <- as.matrix(temp[e+1,1:18])
  i <- tolower(paste(f, g, h, sep=" "))
  
  # expected pattern
  colpattern <- c("na na na",
                  "12 month change existing supply", 
                  "na new build", 
                  "na converted in" ,    
                  "na room additions",
                  "na na closed", 
                  "na converted out", 
                  "na rooms removed",
                  "na gain / loss",
                  "na existing supply",
                  "60 month change existing supply",
                  "na new build",                   
                  "na converted in",
                  "na room additions",
                  "na na closed",
                  "na converted out",
                  "na rooms removed",
                  "na gain / loss")
  # vector to be used for column names
  newnames <- c("segment",
                "pryearsupply", 
                "newbuild", 
                "convertin" ,    
                "rmadditions",
                "rmclosures", 
                "convertout", 
                "rmsremoved",
                "gainloss",
                "supe",
                "fiveyrprsupply",
                "fiveyrnewbuild",                   
                "fiveyrconvertin",
                "fiveyrrmadditions",
                "fiveyrclosures",
                "fiveyrconvertout",
                "fiveyrrmsremoved",
                "fiveyrgainloss",
                "X19",
                "sourcemonth")
  
  # checks whether the two vectors are identical
  # and if so, use the new names to rename the columns
  if (identical(i,colpattern)) {
    colnames(temp) <- newnames
  }
  k <- e+1
  temp <- temp[-(1:k), ]  
  temp$X19 <- NULL
  
  # get ready to filter
  temp$segment <- gsub(" ", "", temp$segment) 
    
  target <- c("LuxuryTotal", "UpperUpscaleTotal", "UpscaleTotal", "UpperMidscaleTotal", "MidscaleTotal", 
              "MidscalewithF&BTotal", "Midscalew/outF&BTotal", "EconomyTotal",
              "IndependentsTotal")
  target <- c(target, "Luxury", "UpperUpscale", "Upscale", "UpperMidscale", "Midscale", 
              "MidscalewithF&B", "Midscalew/outF&B", "Economy")
  backup <- temp
  temp <- backup
  temp <- temp %>%
    filter(segment %in% target) %>%
    filter(! is.na(pryearsupply))%>%
    mutate(segment = gsub("Total", "", segment)) %>%
    mutate(segment = gsub("Luxury", "luxus", segment)) %>%
    mutate(segment = gsub("UpperUpscale", "upuus", segment)) %>%
    mutate(segment = gsub("Upscale", "upsus", segment)) %>%
    mutate(segment = gsub("UpperMidscale", "upmus", segment)) %>%
    mutate(segment = gsub("MidscalewithF&B", "midwus", segment)) %>%
    mutate(segment = gsub("Midscalew/outF&B", "midwous", segment)) %>%
    mutate(segment = gsub("Midscale", "midus", segment)) %>%
    mutate(segment = gsub("Economy", "ecous", segment)) %>%
    mutate(segment = gsub("Independents", "indus", segment)) %>%
    select(sourcemonth, segment, pryearsupply:fiveyrgainloss)
  
  # removes commas in columns 4 to 13
  temp[, 3:19] <- apply(temp[, 3:19], 2, function(x) gsub(",", "", x))
  temp[, 3:19] <- apply(temp[, 3:19], 2, function(x) gsub("\\(", "-", x))
  temp[, 3:19] <- apply(temp[, 3:19], 2, function(x) gsub("\\)", "", x))
  temp[, 3:19] <- apply(temp[, 3:19], 2, function(x) as.numeric(x))
  
  head(temp)
  out_changes <- rbind(out_changes, temp)
}

# adds US totals, summing all entries for a given month
ustot_changes <- out_changes %>%
  group_by(sourcemonth) %>% 
  summarise_each(funs(sum), pryearsupply:fiveyrgainloss) %>%
  mutate(segment="ustot")
out_changes <- rbind(out_changes, ustot_changes)  

# adds column that is the year of the changes
# this simplifies a bit, because in many cases I have changes through
# November, because that was the pipeline info that was in the December
# summary reports released prior to December 2013. After that, they started
# going with December census rather than November
out_changes <- out_changes %>%
  mutate(yrofchange=as.Date(paste(year(sourcemonth), "-01-01", sep=""))) %>%
  select(sourcemonth, yrofchange,segment, pryearsupply:fiveyrgainloss) %>%
  arrange(segment, sourcemonth)

# saves Rdata version of the data
save(out_changes, file="output_data/out_changes.Rdata")