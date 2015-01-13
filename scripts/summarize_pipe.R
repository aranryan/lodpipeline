

load("output_data/out_open_year.Rdata")
load("output_data/out_changes.Rdata")

out_open_year <- out_open_year %>%
  filter(segment =="ustot")%>%
  arrange(hori, sourcemonth)

out_changes <- out_changes %>%
  #filter(segment =="ustot")%>%
  select(sourcemonth:segment, newbuild)

out_changes <- out_changes %>%
  select(-sourcemonth) 

open_year <- left_join(out_open_year, out_changes, by = c("open_year" = "yrofchange", "segment" = "segment"))

open_year <- open_year %>%
  select(sourcemonth, open_year, segment, hori, horitext, newbuild, rmsincns:htlsuncnf)

write.csv(open_year, file="output_data/open_year.csv")
