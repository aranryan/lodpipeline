
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


load("output_data/out_open_year.Rdata")
scalstat <- readr::read_csv(file="input_data/scalstat.csv", 
                     col_names = TRUE, 
                     cols(.default = "n", 
                          date = "c", 
                          scale = "c",
                          type = "c")) %>%
  mutate(date = as.Date(date))

# load("output_data/out_changes.Rdata")
# this was previously using the out_changes file as a source of the 
# actual newbuild. I've changed it to use data from the scalstat file. 

# out_open_year <- out_open_year %>%
#   #filter(segment =="ustot")%>%
#   arrange(hori, sourcemonth)
# 
# out_changes <- out_changes %>%
#   #filter(segment =="ustot")%>%
#   select(sourcemonth:segment, newbuild)
# 
# out_changes <- out_changes %>%
#   # filter to just Nov and Dec source months, as this will give us one reading on each year
#   filter(srcm_month %in% c("Nov", "Dec")) %>%
#   # then drop the source month fields
#   select(-sourcemonth, -srcm_month)

# convert the scalstat data to annual
scalstat_a <- scalstat %>%
  mutate(year = lubridate::year(date)) %>%
  filter(type %in% c("rooms")) %>%
  group_by(year, scale) %>%
  summarize_if(is.numeric, sum) %>%
  ungroup() %>%
  mutate(date = as.Date(paste0(year, "-01-01"))) %>%
  select(-year) %>%
  select(date, scale, everything())

open_year_1 <- left_join(out_open_year, scalstat_a, 
                       by = c("open_year" = "date", "segment" = "scale"))

open_year_2 <- open_year_1 %>%
  select(sourcemonth, srcm_month, open_year, segment, hori, horitext, rmsincns, 
         rmsfnlpl, rmsplnnn, rmstotal, open, close, adds, drop, cin, cout, netchg)

open_year <- open_year_2

write_csv(open_year, path="output_data/open_year.csv")
