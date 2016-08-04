

load("output_data/out_open_year.Rdata")
load("output_data/out_changes.Rdata")

write.csv(out_changes, file="output_data/out_changes.csv", row.names = FALSE)

out_open_year <- out_open_year %>%
  #filter(segment =="ustot")%>%
  arrange(hori, sourcemonth)

out_changes <- out_changes %>%
  #filter(segment =="ustot")%>%
  select(sourcemonth:segment, newbuild)

out_changes <- out_changes %>%
  # filter to just Nov and Dec source months, as this will give us one reading on each year
  filter(srcm_month %in% c("Nov", "Dec")) %>%
  # then drop the source month fields
  select(-sourcemonth, -srcm_month)

open_year <- left_join(out_open_year, out_changes, 
                       by = c("open_year" = "yrofchange", "segment" = "segment"))

open_year <- open_year %>%
  select(sourcemonth, srcm_month, open_year, segment, hori, horitext, newbuild, rmsincns:htlsuncnf)

write.csv(open_year, file="output_data/open_year.csv", row.names = FALSE)
