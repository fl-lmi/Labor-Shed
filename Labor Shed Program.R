#----PART A: TCOB

#A1-clear global environment
rm(list=ls())

#A2-packages
library(rvest)
library(RODBC)
library(tidyverse)
library(reshape2)
library(ReporteRs)
library(scales)
library(dplyr)
library(openxlsx)

#A3-identify counties of interest
area<-'109'                     #<-<-<-ENTER FIPS CODES OF STUDY AREA HERE!!! 
#      (will handle single county or combination of counties) 

area_name <- 'St Johns County'  #<-<-<-ENTER STUDY AREA NAME HERE!!!

html_area<-'stjohnscountyflorida' #<-<-<-ENTER QUICKFACTS URL AREA HERE!!!

pop_data_range <- 2011:2017

report_date <- 'May 2018'  #<-<-<-ENTER MONTH AND YEAR OF REPORT HERE!!!

pop_file_name <- 'PEP_2017_PEPANNRES.csv'

######LAUS/POPULATION TABLE PARAMETERS#################
wfr<-c('000008')                                      #<-<-<-ENTER WORKFORCE REGION NUMBER HERE!!!
month<-'02'                                           #<-<-<-ENTER MONTH NUMBER HERE!!!
year<-2018                                            #<-<-<-ENTER CURENT YEAR HERE!!! 
month_name<-'February'                                #<-<-<-ENTER MONTH NAME HERE!!!
################DERIVED GLOBAL PARAMETERS##############
prev_year<-year-1                                     # 
month_year<-paste(month_name,year,sep = ' ')          #
month_prev_year<-paste(month_name,prev_year,sep = ' ')#
#######################################################


#set environment options
options(scipen = 999) #<-no scientific notation

#----PART B: READ IN (AND FORMAT) FLORIDA WAC FILE

#B1a-set working directory

setwd('T:/Census Data Center/On-The-Map/Automation/WAC-OD-RAC')


#B1b-read WAC into environment

wac<-read.csv('fl_wac_S000_JT01_2015.csv'
              ,header = TRUE
              ,colClasses = c('character',rep('numeric',51),'character'))



#----PART C: READ IN (AND FORMAT) FLORIDA RAC FILE


#C1-read RAC into environment

rac<- read.csv('fl_rac_S000_JT01_2015.csv'
               ,header = TRUE
               ,colClasses = c('character',rep('numeric',41),'character'))


#----PART D: READ IN (AND FORMAT) FLORIDA OD FILES AND OUTFLOW RANKINGS

#D1-read OD into environment

od_main <- read.csv('fl_od_main_JT01_2015.csv'
                    ,header = TRUE
                    ,colClasses = c(rep('character',2),rep('numeric',10),'character'))

od_aux <- read.csv('fl_od_aux_JT01_2015.csv'
                   ,header = TRUE
                   ,colClasses = c(rep('character',2),rep('numeric',10),'character'))

od <- bind_rows(od_main,od_aux)


#D2-read county fips codes into environment 


fl_county_fips <- read.table('FLCountyFIPS.txt', sep = ',', header = TRUE, colClasses = 'character')
us_county_fips <- read.table('USCountyFIPS.txt', sep = ',', header = TRUE, colClasses = 'character')


#D3-read outflow rankings

outflow_rank <- read.xlsx('T:/Census Data Center/On-The-Map/Outflow by RWB and County.xlsx', sheet = 2, startRow=2)
outflow_rank <- outflow_rank[, c(1,(length(outflow_rank)-1),length(outflow_rank))]
colnames(outflow_rank)[c(1,3)] <- c('LaborOutflow','HighestOutflow')
outflow_rank <- arrange(outflow_rank, HighestOutflow)

#----PART E: CREATE STUDY TABLES FOR WAC, RAC, AND OD, RESPECTIVELY

#   Parse w_geocode and h_geocode into state, county, and sub-county components.
#   Filter so that resulting table only contains study area observations.

wac %>%
  mutate(st = substr(w_geocode,1,2)   #<-state fips
         ,co = substr(w_geocode,3,5)   #<-county fips
         ,subco = substr(w_geocode,6,15)   #<-subcounty fips
         ,study = ifelse(co %in% area,1,0)   #<-study area marker
  ) %>%
  #filter by study marker
  filter(study == 1
         #store results in new table called study_wac       
  ) -> study_wac


rac %>% 
  mutate( h_st = substr(h_geocode,1,2)  #<-state fips
          ,h_co = substr(h_geocode,3,5)   #<-county fips
          ,h_subco = substr(h_geocode,6,15) #<-subcounty fips
  ) %>%
  filter(h_co %in% area
         #store results in new table called study_rac       
  ) -> study_rac


od %>% 
  mutate(w_st = substr(w_geocode,1,2)
         ,h_st = substr(h_geocode,1,2)  #<-state fips
         ,w_co = substr(w_geocode,3,5)
         ,h_co = substr(h_geocode,3,5)   #<-county fips
         ,w_subco = substr(w_geocode,6,15)
         ,h_subco = substr(h_geocode,6,15) #<-subcounty fips
  ) %>%
  filter(w_co %in% area | h_co %in% area
         #store results in new table called study_od       
  ) -> study_od


#verify that the specified study area fips codes are the only values in "co".

table(study_wac$co)
table(study_rac$h_co)
table(study_od$h_co)
table(study_od$h_co)


#----PART F: MELT STUDY TABLES 

#F1a-melt study_rac table

study_wac %>%
  melt(id.vars = c('w_geocode','createdate','st','co','subco','study')) %>%
  #order by geocode and variable
  arrange(w_geocode,variable
          #store results in new table called study_wac_melt
  ) -> study_wac_melt


#F1b-Verify successful melt. Create melt_id composed of geocode and variable type. There should
#    be no duplicates of melt_id; results of table command should all be FALSE.

melt_wac_id<-paste0(study_wac_melt$w_geocode,'_',study_wac_melt$variable)
table(duplicated(melt_wac_id))


#F1c-Create separate vectors for age, earnings, industry, race, ethnicity, education, and sex variable names.
wac_age_vars<-c('CA01','CA02','CA03')   #<-age variables
wac_earn_vars<-c('CE01','CE02','CE03')   #<-earnings variables
wac_ind_vars<-c('CNS01','CNS02','CNS03','CNS04','CNS05','CNS06','CNS07','CNS08','CNS09','CNS10'    #<-industry variables
            ,'CNS11','CNS12','CNS13','CNS14','CNS15','CNS16','CNS17','CNS18','CNS19','CNS20')   #<-industry variables
wac_race_vars<-c('CR01','CR02','CR03','CR04','CR05','CR07')   #<-race variables
wac_eth_vars<-c('CT01','CT02')   #<-ethnicity variables
wac_edu_vars<-c('CD01','CD02','CD03','CD04')   #<-education variables
wac_sex_vars<-c('CS01','CS02')   #<-sex variables


#F1d-Define new "type" variable that indicates if a variable describes age, race, education, etc.
study_wac_melt %>%
  mutate(type = 'other'
         ,type = ifelse(variable == 'C000','all',type)
         ,type = ifelse(variable %in% wac_age_vars,'age',type)
         ,type = ifelse(variable %in% wac_earn_vars,'earn',type)
         ,type = ifelse(variable %in% wac_ind_vars,'ind',type)
         ,type = ifelse(variable %in% wac_race_vars,'race',type)
         ,type = ifelse(variable %in% wac_eth_vars,'eth',type)
         ,type = ifelse(variable %in% wac_edu_vars,'edu',type)
         ,type = ifelse(variable %in% wac_sex_vars,'sex',type)
  ) -> study_wac_melt


#F2a-melt study_rac table

study_rac %>%
  melt(id.vars = c('h_geocode', 'createdate','h_st','h_co','h_subco')) %>%
  #order by geocode and variable
  arrange(h_geocode,variable
          #store results in new table called study_rac_melt
  ) -> study_rac_melt


#F2b-Verify successful melt. Create melt_id composed of geocode and variable type. There should
#    be no duplicates of melt_id; results of table command should all be FALSE.

melt_rac_id<-paste0(study_rac_melt$w_geocode,'_',study_rac_melt$variable)
table(duplicated(melt_rac_id))


#F2c-Create separate vectors for age, earnings, industry, race, ethnicity, education, and sex variable names.

rac_age_vars<-c('CA01','CA02','CA03')   #<-age variables
rac_earn_vars<-c('CE01','CE02','CE03')   #<-earnings variables
rac_ind_vars<-c('CNS01','CNS02','CNS03','CNS04','CNS05','CNS06','CNS07','CNS08','CNS09','CNS10'    #<-industry variables
            ,'CNS11','CNS12','CNS13','CNS14','CNS15','CNS16','CNS17','CNS18','CNS19','CNS20')   #<-industry variables
rac_race_vars<-c('CR01','CR02','CR03','CR04','CR05','CR07')   #<-race variables
rac_eth_vars<-c('CT01','CT02')   #<-ethnicity variables
rac_edu_vars<-c('CD01','CD02','CD03','CD04')   #<-education variables
rac_sex_vars<-c('CS01','CS02')   #<-sex variables


#F2d-Define new "type" variable that indicates if a variable describes age, race, education, etc.

study_rac_melt %>%
  mutate(type = 'other'
         ,type = ifelse(variable == 'C000','all',type)
         ,type = ifelse(variable %in% rac_age_vars,'age',type)
         ,type = ifelse(variable %in% rac_earn_vars,'earn',type)
         ,type = ifelse(variable %in% rac_ind_vars,'ind',type)
         ,type = ifelse(variable %in% rac_race_vars,'race',type)
         ,type = ifelse(variable %in% rac_eth_vars,'eth',type)
         ,type = ifelse(variable %in% rac_edu_vars,'edu',type)
         ,type = ifelse(variable %in% rac_sex_vars,'sex',type)
  ) -> study_rac_melt

#F3a-melt study_od table

study_od %>%
  melt(id.vars = c('w_geocode', 'h_geocode', 'createdate','w_st', 'h_st', 'w_co','h_co','w_subco','h_subco')) %>%
  #order by geocode and variable
  arrange(w_geocode, h_geocode, variable
          #store results in new table called study_od_melt
  ) -> study_od_melt


#F3b-Verify successful melt. Create melt_id composed of geocode and variable type. There should
#    be no duplicates of melt_id; results of table command should all be FALSE.

melt_od_id<-paste0(study_od_melt$w_geocode, study_od_melt$h_geocode, study_od_melt$variable, sep = '_')
table(duplicated(melt_od_id))

#F3c-Create separate vectors for age, earnings, industry, race, ethnicity, education, and sex variable names.

od_age_vars<-c('SA01','SA02','SA03')   #<-age variables
od_earn_vars<-c('SE01','SE02','SE03')   #<-earnings variables
od_ind_vars<-c('SI01','SI02','SI03')   #<-industry variables


#F3d-Define new "type" variable that indicates if a variable describes age, race, education, etc.

study_od_melt %>%
  mutate(type = 'other'
         ,type = ifelse(variable == 'S000','all',type)
         ,type = ifelse(variable %in% od_age_vars,'age',type)
         ,type = ifelse(variable %in% od_earn_vars,'earn',type)
         ,type = ifelse(variable %in% od_ind_vars,'ind',type)
  ) -> study_od_melt



#---PART G: CREATE WAC SUMMARY TABLE

#G1a-Summarize by variable
study_wac_melt %>%
  group_by(variable,type) %>%
  summarise(count = sum(value)#) %>%
  #remove observations of type 'other'
  #filter(type != 'other'
  ) -> study_wac_sum


#G2-Calculate share variable

#G2a-identify total jobs in study area

total_wac_jobs<-study_wac_sum$count[study_wac_sum$variable=='C000']


#G2b-divide all observations by total_wac_jobs to calculate share

study_wac_sum %>%
  mutate(share = round((count/total_wac_jobs)*100,1)
         #store results in study_sum table
  ) -> study_wac_sum


#G2c-Calulate value for "educational attainment not available"

study_wac_sum %>%
  filter(type == 'edu') %>%
  group_by(type) %>%
  summarise(count = sum(count)) %>%
  mutate(count = total_wac_jobs - count
         ,share = round((count/total_wac_jobs)*100,1)
         ,variable = 'CD05'
  ) %>%
  #order columns to match study_wac_sum
  select(variable,type,count,share
         #store results in a one-row table called "CD05_row"
  )-> CD05_wac_row

#confirm CD05_row is a single row

nrow(CD05_wac_row)

#join CD05_row to study_wac_sum

study_wac_sum %>%
  rbind.data.frame(CD05_wac_row
                   #store in a new table called "study all"                 
  ) -> study_wac_all

#confirm study all is one row longer than study_wac_sum

nrow(study_wac_sum)
nrow(study_wac_all)

#G3a-match observations to titles

wac_titles<-read.csv('wac titles.csv',header = TRUE)

#G3b-merge study_wac_all with titles

study_wac_all %>%
  merge(wac_titles
        ,by = 'variable'
        #store results in a new table called "wac_table"
  ) -> wac_table


#confirm wac_table has the same number of observations as study_wac_all

nrow(study_wac_all)
nrow(wac_table)

#select and arrange rows for final table

wac_table %>%
  arrange(order) %>%
  select(title,count,share
  ) -> wac_table_final


#---PART H: CREATE RAC SUMMARY TABLE

#H1a-Summarize by variable

study_rac_melt %>%
  group_by(variable,type) %>%
  summarise(count = sum(value)) %>%
  #remove observations of type 'other'
  filter(type != 'other'
  ) -> study_rac_sum


#H2-Calculate share variable

#H2a-identify total jobs in study area

total_rac_jobs<-study_rac_sum$count[study_rac_sum$variable=='C000']

#H2b-divide all observations by total_jobs to calculate share

study_rac_sum %>%
  mutate(share = round((count/total_rac_jobs)*100,1)
         #store results in study_sum table
  ) -> study_rac_sum


#H2c-Calulate value for "educational attainment not available" and add to study table

study_rac_sum %>%
  filter(type == 'edu') %>%
  group_by(type) %>%
  summarise(count = sum(count)) %>%
  mutate(count = total_rac_jobs - count
         ,share = round((count/total_rac_jobs)*100,1)
         ,variable = 'CD05'
  ) %>%
  #order columns to match study_sum
  select(variable,type,count,share
         #store results in a one-row table called "CD05_row"
  )-> CD05_rac_row

#confirm CD05_row is a single row

nrow(CD05_rac_row)

#join CD05_row to study_sum

study_rac_sum %>%
  bind_rows(CD05_rac_row
            #store in a new table called "study all"                 
  ) -> study_rac_all

#confirm study all is one row longer than study_rac_sum

nrow(study_rac_sum)
nrow(study_rac_all)


#H3a-match observations to titles


rac_titles<-read.csv('T:/Census Data Center/On-The-Map/Automation/WAC/wac titles.csv',header = TRUE)


#H3b-merge study_rac_all with titles
study_rac_all %>%
  merge(rac_titles
        ,by = 'variable'
        #store results in a new table called "pretty_rac_table"
  ) -> rac_table

#confirm pretty_table has the same number of observations as study_rac_all

nrow(study_rac_all)
nrow(rac_table)

#select and arrange rows for final table
rac_table %>%
  arrange(order) %>%
  select(title,count,share
  ) -> rac_table_final



#---PART I: CREATE SUMMARY TABLES FOR THOSE LIVING OR WORKING IN AREA


study_od_melt %>%
  filter(w_co %in% area) %>%
  group_by(variable,type) %>%
  summarise(count = sum(value))  -> study_sum_working # Summary of everyone working in the area


study_od_melt %>%
  filter(h_co %in% area & w_co %in% area & h_st == 12) %>%
  group_by(variable,type) %>%
  summarise(count = sum(value))  -> stud_sum_work_live # Summary of everyone working and living in the area


study_od_melt %>%
  filter((h_st!=12 & w_co %in% area)|(h_st==12 & !(h_co %in% area) & w_co %in% area)) %>%
  group_by(variable,type) %>%
  summarise(count = sum(value))  -> stud_sum_live_external # Summary of everyone who works in the area but lives outside


study_od_melt %>%
  filter((h_co %in% area) & (h_st==12)) %>%
  group_by(variable, w_st, w_co, type) %>%
  summarise(count = sum(value)) %>%
  arrange(variable, desc(count)) -> stud_sum_work_counties # Summary of where workers living in area are employed


study_od_melt %>%
  filter(w_co %in% area) %>%
  group_by(variable, h_st, h_co, type) %>%
  summarise(count = sum(value)) %>%
  arrange(variable, desc(count)) -> stud_sum_live_counties # Summary of where workers employed in area live


#---PART J: CALCULATE OUTFLOW JOB CHARACTERISTICS


#Industry Class variables 

goods <- c('CNS01', 'CNS02', 'CNS04', 'CNS05') #Vector of variables for Goods Producing Industry Class


trade <- c('CNS06', 'CNS07', 'CNS08', 'CNS03') # Vector of variables for Trade, Transportation, and Utilities Class


aos <- c('CNS09') 

for (i in 10:20) {
  aos <- append(aos, paste('CNS', i, sep=""))   #Vector of variables for All Other Services Industry Class
}

# Selection Area Labor Market Size

area_working <- study_sum_working$count[study_sum_working$variable=='S000'] #object for # working in area
 

area_living <-  rac_table_final$count[rac_table_final$title == 'Total Primary Jobs'] #object for # living in area
                                                                                                    #Requires RAC to be run

net_flow <- area_working - area_living  #object for net inflow or outflow


#In-Area Labor Force Efficiency

area_live_work <- stud_sum_work_live$count[stud_sum_work_live$variable=='S000'] #object for # living and working in area


area_live_out_work <- area_living - area_live_work  #object for # living in area but working outside


#In-Area Employment Efficiency (some overlap with previous sections)

area_work_out_live <- area_working - area_live_work #object for # working in area but living outside


#Outflow age characteristics

outflow_lt29 <- rac_table_final$count[rac_table_final$title == 'Age 29 or younger'] -
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SA01']


outflow_30_54 <- rac_table_final$count[rac_table_final$title == 'Age 30 to 54'] -
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SA02']


outflow_gt55 <- rac_table_final$count[rac_table_final$title == 'Age 55 or older'] -
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SA03']


#Outflow employment characteristics

outflow_earn_lt1250 <- rac_table_final$count[rac_table_final$title == '$1,250 per month or less'] -
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SE01']

outflow_earn_1251 <-rac_table_final$count[rac_table_final$title == '$1,251 to $3,333 per month'] -
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SE02']

outflow_earn_mt3333 <- rac_table_final$count[rac_table_final$title == 'More than $3,333 per month'] -
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SE03']


#Outflow industry characteristics 

outflow_goods <- sum(study_rac_all$count[which(study_rac_all$variable %in% goods)]) - 
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SI01']

outflow_trade <- sum(study_rac_all$count[which(study_rac_all$variable %in% trade)]) - 
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SI02']

outflow_aos <- sum(study_rac_all$count[which(study_rac_all$variable %in% aos)]) - 
  stud_sum_work_live$count[stud_sum_work_live$variable == 'SI03']


#---PART I: CONSTRUCT INFLOW/OUTFLOW TABLE


#Titles for Inflow/Outflow Table

in_out_titles <- data.frame(c('Employed in the Selection Area'
                   , 'Living in the Selection Area'
                   , 'Net Job Inflow (+) or Outflow (-)'
                   , 'Living in the Selection Area'
                   , 'Living and Employed in the Selection Area'
                   , 'Living in the Selection Area but Employed Outside'
                   , 'Employed in the Selection Area'
                   , 'Employed and Living in the Selection Area'
                   , 'Employed in the Selection Area but Living Outside'
                   , 'External Jobs Filled by Residents'
                   , 'Workers Aged 29 or younger'
                   , 'Workers Aged 30 to 54'
                   , 'Workers Aged 55 or older'
                   , 'Workers Earning $1,250 per month or less'
                   , 'Workers Earning $1,251 to $3,333 per month'
                   , 'Workers Earning More than $3,333 per month'
                   , 'Workers in the "Goods Producing" Industry Class'
                   , 'Workers in the "Trade, Transportation, and Utilities" Industry Class'
                   , 'Workers in the "All Other Services" Industry Class'
                   , 'Internal Jobs Filled by Outside Workers'
                   , 'Workers Aged 29 or younger'
                   , 'Workers Aged 30 to 54'
                   , 'Workers Aged 55 or older'
                   , 'Workers Earning $1,250 per month or less'
                   , 'Workers Earning $1,251 to $3,333 per month'
                   , 'Workers Earning More than $3,333 per month'
                   , 'Workers in the "Goods Producing" Industry Class'
                   , 'Workers in the "Trade, Transportation, and Utilities" Industry Class'
                   , 'Workers in the "All Other Services" Industry Class'
                   , 'Internal Jobs Filled by Residents'
                   , 'Workers Aged 29 or younger'
                   , 'Workers Aged 30 to 54'
                   , 'Workers Aged 55 or older'
                   , 'Workers Earning $1,250 per month or less'
                   , 'Workers Earning $1,251 to $3,333 per month'
                   , 'Workers Earning More than $3,333 per month'
                   , 'Workers in the "Goods Producing" Industry Class'
                   , 'Workers in the "Trade, Transportation, and Utilities" Industry Class'
                   , 'Workers in the "All Other Services" Industry Class'))

colnames(in_out_titles) <- c('Title')


# Data Frame of manually calculated characteristics

in_out_calc <- data.frame(count = c(area_working
                                    ,area_living
                                    ,net_flow
                                    ,area_living
                                    ,area_live_work
                                    ,area_live_out_work
                                    ,area_working
                                    ,area_live_work
                                    ,area_work_out_live
                                    ,area_live_out_work
                                    ,outflow_lt29
                                    ,outflow_30_54
                                    ,outflow_gt55
                                    ,outflow_earn_lt1250
                                    ,outflow_earn_1251
                                    ,outflow_earn_mt3333
                                    ,outflow_goods
                                    ,outflow_trade
                                    ,outflow_aos))

# Data frame of characteristics for outside workers

stud_sum_live_external_count <- data.frame(stud_sum_live_external$count)
colnames(stud_sum_live_external_count) <- c('count')


# Data frame of characteristics for residents

stud_sum_work_live_count <- data.frame(stud_sum_work_live$count)
colnames(stud_sum_work_live_count) <- c('count')


# Data frame of all characteristics (combines manually calculated, outside workers, and residents data frames)

in_out_numbers <- bind_rows(in_out_calc, stud_sum_live_external_count, stud_sum_work_live_count)


# Calculate shares

in_out_numbers <- mutate(in_out_numbers, Share = NA)


in_out_numbers$Share[c(1:2,7:9)] <- map_dbl(in_out_numbers$count[c(1:2,7:9)], function(x) round((x/area_working)*100,1))

in_out_numbers$Share[4:6] <- map_dbl(in_out_numbers$count[4:6], function(x) round((x/area_living)*100,1))

in_out_numbers$Share[10:19] <- map_dbl(in_out_numbers$count[10:19], 
                                       function(x) round((x/area_live_out_work)*100,1))

in_out_numbers$Share[20:29] <- map_dbl(in_out_numbers$count[20:29], 
                                       function(x) round((x/stud_sum_live_external$count[stud_sum_live_external$variable=='S000'])*100,1))

in_out_numbers$Share[30:39] <- map_dbl(in_out_numbers$count[30:39], 
                                       function(x) round((x/stud_sum_work_live$count[stud_sum_work_live$variable=='S000'])*100,1))



# Data frame of Inflow/Outflow Characteristics

in_out_summ <- bind_cols(in_out_titles,in_out_numbers) #Final table for inflow/outflow characteristics


#---PART K: CONSTRUCT WORK DESTINATION SUMMARY

stud_sum_work_counties %>% #
  .[1:25, c('w_co','count')] %>%
  merge(fl_county_fips,., by.y = 'w_co', by.x = 'County_FIPS') %>%
  mutate(title=paste0(County,sep=', ',State)) %>%
  select(title,count) %>%
  arrange(desc(count)) -> work_dest


work_dest <- bind_rows(rac_table_final[1,1:2], work_dest)

work_des_other <- data.frame(
  title = 'All Other Locations', count= work_dest$count[1]-sum(work_dest$count[2:length(work_dest$count)]))

work_dest <- bind_rows(work_dest,work_des_other)

work_dest %>% 
  mutate(share = round((count/count[1])*100,1)) -> work_dest_final #final table for work-destination 



#---PART L: CONSTRUCT HOME DESTINATION SUMMARY

us_county_fips %>%
  mutate(match_id = paste0(State_FIPS,'_',County_FIPS)) -> us_county_fips


stud_sum_live_counties %>% #
  mutate(match_id = paste0(h_st,'_',h_co)) %>%
  .[1:25, c('match_id','count')] %>%
  merge(us_county_fips, ., by='match_id') %>%
  mutate(title=paste0(County,sep=', ',State)) %>%
  select(title,count) %>%
  arrange(desc(count)) -> home_dest


home_dest <- bind_rows(wac_table_final[1,1:2], home_dest)  #Requires WAC table

home_des_other <- data.frame(
  title = 'All Other Locations', count= home_dest$count[1]-sum(home_dest$count[2:length(home_dest$count)]))

home_dest <- bind_rows(home_dest,home_des_other)

home_dest %>% 
  mutate(share = round((count/count[1])*100,1)) -> home_dest_final #final table for home-destination




#---PART M: FORMAT NUMBERS OF ALL TABLES BEFORE MOVING INTO WORD DOCUMENT


wac_table_final$share <- sprintf("%.1f",wac_table_final$share)
wac_table_final$count <- comma(wac_table_final$count)
wac_table_final$share <- paste0(wac_table_final$share, '%', sep="")


rac_table_final$count <- comma(rac_table_final$count)
rac_table_final$share <- sprintf("%.1f",rac_table_final$share)
rac_table_final$share <- paste0(rac_table_final$share, '%', sep="")

in_out_summ$count <- comma(in_out_summ$count)
in_out_summ$Share <- sprintf("%.1f", in_out_summ$Share)
in_out_summ$Share <- paste0(in_out_summ$Share,'%',sep="")
in_out_summ$Share[in_out_summ$Title=='Net Job Inflow (+) or Outflow (-)'] <- '-'

work_dest_final$count <- comma(work_dest_final$count)
work_dest_final$share <- sprintf("%.1f",work_dest_final$share)
work_dest_final$share <- paste0(work_dest_final$share, '%', sep="")

home_dest_final$count <- comma(home_dest_final$count)
home_dest_final$share <- sprintf("%.1f",home_dest_final$share)
home_dest_final$share <- paste0(home_dest_final$share, '%', sep="")



############################PART N: SCRAPE CENSUS QUICKFACTS TABLE##############################

#Scrape------------------------------------------------------------
qfacts_url<-paste0('https://www.census.gov/quickfacts/fact/table/',
                   html_area,
                   ',FL,US/PST045216')

#identify census quickfacts url 
qfacts_page<-read_html(qfacts_url)

#scrape variable names column
qfacts_page %>%
  html_nodes('td:nth-child(1)') %>%
  html_text() %>%
  #remove non-ASCII characters in areaname
  iconv('latin1','ASCII',sub = '') %>%
  gsub('\n','',.) %>%
  gsub(',  \\(V2016\\)','',.)-> vars



#scrape county data column
qfacts_page %>%
  html_nodes('td:nth-child(2)') %>%
  html_text() %>%
  iconv('latin1','ASCII',sub = '') %>%
  gsub('\n','',.) -> county

#scrape state data column
qfacts_page %>%
  html_nodes('td:nth-child(3)') %>%
  html_text() %>%
  iconv('latin1','ASCII',sub = '') %>%
  gsub('\n','',.) -> state

#scrape national data column
qfacts_page %>%
  html_nodes('td:nth-child(4)') %>%
  html_text() %>%
  iconv('latin1','ASCII',sub = '') %>%
  gsub('\n','',.) -> nation

#compile all columns into data frame
qfacts <- data.frame(vars,county,state,nation)

#Divide up quickfacts data subject tables------------------------------------------------
#header
qf_header<-data.frame('',area_name,'Florida','United States')

#page one table
pop_header<-as.data.frame('Population')
colnames(pop_header) <- 'vars'
qf_pop_table <- bind_rows(pop_header,
                        qfacts[2:8,])

#pop_table<-as.data.frame(qfacts[2:8,])

#age and sex table
age_header<-as.data.frame('Age and Sex')
colnames(age_header) <- 'vars'
age_table <- bind_rows(age_header,
                        qfacts[9:12,])

#race table
race_header<-as.data.frame('Race and Hispanic Origin')
colnames(race_header) <- 'vars'
race_table <- bind_rows(race_header,
                         qfacts[13:20,])

#population characteristics table
char_header<-as.data.frame('Population Characteristics')
colnames(char_header) <- 'vars'
char_table <- bind_rows(char_header,
                         qfacts[21:22,])

#housing table
hous_header<-as.data.frame('Housing')
colnames(hous_header) <- 'vars'
hous_table <- bind_rows(hous_header,
                         qfacts[23:29,])

#families and living arrangements table
fam_header<-as.data.frame('Families & Living Arrangements')
colnames(fam_header) <- 'vars'
fam_table <- bind_rows(fam_header,
                        qfacts[30:33,])

#education table
edu_header<-as.data.frame('Education')
colnames(edu_header) <- 'vars'
edu_table <- bind_rows(edu_header,
                        qfacts[34:35,])

#health table
heal_header<-as.data.frame('Health')
colnames(heal_header) <- 'vars'
heal_table <- bind_rows(heal_header,
                         qfacts[36:37,])

#economy table
econ_header<-as.data.frame('Economy')
colnames(econ_header) <- 'vars'
econ_table <- bind_rows(econ_header,
                         qfacts[38:45,])

#transportation table
tran_header<-as.data.frame('Transportation')
colnames(tran_header) <- 'vars'
tran_table <- bind_rows(tran_header,
                         qfacts[46,])

#income and poverty table
inc_header<-as.data.frame('Income & Poverty')
colnames(inc_header) <- 'vars'
inc_table <- bind_rows(inc_header,
                        qfacts[47:49,])

#business table
bus_header<-as.data.frame('Business')
colnames(bus_header) <- 'vars'
bus_table <- bind_rows(bus_header,
                        qfacts[50:61,])
bus_table$state<-as.character(bus_table$state)
bus_table$state[2:8]<-substr(bus_table$state[2:8],1,nchar(bus_table$state[2:8])-1)

#geography table
geo_header<-as.data.frame('Geography')
colnames(geo_header) <- 'vars'
geo_table <- bind_rows(geo_header,
                        qfacts[62:64,])

#create legend table
legend<-rbind.data.frame('Legend'
                         ,'The vintage year (e.g., V2016) refers to the final year of the series (2010 thru 2016). Different vintage years of estimates are not comparable.'
                         ,'Fact Notes'
                         ,'(a) Includes persons reporting only one race'
                         ,'(b) Hispanics may be of any race, so also are included in applicable race categories'
                         ,'(c) Economic Census - Puerto Rico data are not comparable to U.S. Economic Census data'
                         ,'Value Flags'
                         ,'"-" No or too few sample observations were available to compute an estimate.'
                         ,'"D" Suppressed to avoid disclosure of confidential information'
                         ,'"F" Fewer than 25 firms'
                         ,'"NA" Not available'
                         ,'"S" Suppressed; does not meet publication standards'
                         ,'"Z" Value greater than zero but less than half unit of measure shown')

###################################PART O: LAUS AND POPULATION TABLES###################################



#----PART B: CONNECT TO WID AND QUERY LAUS AND POPULATION DATA

#B1-establish odbc connection
wid<-odbcConnect('DEOSQLTEST')

#################################################SQL COMMANDS####################################################
#B2-use wid database
use_wid<-sqlQuery(wid, paste("use tstLMS_WID"))

#B3-laus workforce region data
wfr_laus<-sqlQuery(wid, paste0("--SELECT WORKFORCE REGION LAUS DATA
                               select s.areatype, s.area, l.periodyear, l.period, sum(l.laborforce) as 'laborforce', sum(l.emplab) as 'emplab', sum(l.unemp) as 'unemp', (sum(l.unemp)/sum(l.laborforce))*100 as 'unemprate'
                               from labforce l 
                               left join subgeog s on s.stfips = l.stfips and s.subareatyp = l.areatype and s.subarea = l.area
                               where s.areatype = '15'
                               and s.area = ", wfr, 
                               "and l.stfips = '12'
                               and l.areatype = '04'
                               and l.adjusted = '0'
                               and l.periodtype = '03'
                               and l.periodyear in ('",year,"','",prev_year,"')
                               and l.period = ", month,
                               "group by s.area, s.areatype, l.periodyear, l.period"))
#B4-laus county data                          
co_laus<-sqlQuery(wid, paste0("--SELECT COUNTY LAUS DATA
                              select s.subareatyp as 'areatype', s.subarea as 'area', l.periodyear, l.period, l.laborforce, l.emplab, l.unemp, l.unemprate
                              from labforce l 
                              left join subgeog s on s.stfips = l.stfips and s.subareatyp = l.areatype and s.subarea = l.area
                              where s.areatype = '15'
                              and s.area = ", wfr,
                              "and l.stfips = '12'
                              and l.areatype = '04'
                              and l.adjusted = '0'
                              and l.periodtype = '03'
                              and l.periodyear in ('",year,"','",prev_year,"')
                              and l.period =", month))
#B5-laus state and national data                          
fl_us_laus<-sqlQuery(wid, paste0("--SELECT STATE ANS NATIONAL LAUS DATA
                                 select areatype, area, periodyear, period, laborforce, emplab, unemp, unemprate
                                 from labforce
                                 where stfips in ('12','00')
                                 and areatype in ('01','00')
                                 and periodtype = '03'
                                 and periodyear in ('",year,"','",prev_year,"')
                                 and period =", month,
                                 "and adjusted = '0'"))


#B6-geography name data
geo_names<-sqlQuery(wid, paste0("select areatype, area, areaname
                                from geog
                                where stfips in ('12','00')
                                and areatype in ('01','00','15','04')"))

#B7&8-County, state, and national population data 

full_pop_data <- read.csv(pop_file_name, skip=1)
full_pop_data <- full_pop_data[,c(3,6:length(full_pop_data))]
colnames(full_pop_data) <- c('areaname', 
                             paste0('y', pop_data_range))

full_pop_data$areaname <- sub(', Florida', '', full_pop_data$areaname)

wfr_counties <- sqlQuery(wid, paste('select g.areaname from subgeog s, geog g
                                    where s.stfips = 12 
                                    and s.areatype = 15 
                                    and s.area =', wfr,
                                    'and subareatyp = 04
                                    and subarea in (
                                    select area from geog
                                    where stfips = 12
                                    and areatype = 04
                                    )
                                    and s.substfips + s.subareatyp + s.subarea = g.stfips + g.areatype + g.area
                                    order by g.areaname
                                    ', sep = ' '))


co_pop <- filter(full_pop_data, areaname %in% wfr_counties$areaname)

wfr_pop <- data.frame(areaname = geo_names$areaname[geo_names$areatype==15 & geo_names$area == as.double(wfr)],
                      t(map_dbl(co_pop[,-1], sum)))

st_natn_pop <- arrange(filter(full_pop_data, areaname=='Florida'|areaname=='United States'), areaname)

pop_table <- bind_rows(wfr_pop,co_pop,st_natn_pop)
pop_table <- comma(pop_table)


#close odbc connection
odbcClose(wid)
##################################################END OF SQL COMMANDS########################################################

#----PART C: ORGANIZE LAUS TABLE

#C1a-combine workforce, county, and state/national tables
bind_rows(wfr_laus,co_laus,fl_us_laus) %>%
  #C1b-remove period column
  mutate(period = NULL
         #C1c-replace year numbers with generic year names 'current' and 'previous'
         ,periodyear = ifelse(periodyear == year,'current','previous')
  ) -> all_laus

#C2-melt laus table
laus_melt<-melt(all_laus,id.vars = c('areatype','area','periodyear'))

#C3-reform laus table so that each year/variable combination has its own column
laus_cast<-dcast(laus_melt,areatype + area ~ periodyear + variable)

#C4-merge area names to laus table
laus_names<-merge(laus_cast,geo_names,by = c('areatype','area'),all.x = TRUE)

#C5-arrange rows in desired order
area_order<-c(15,4,1,0)
laus_names<-laus_names[order(match(laus_names$areatype,area_order),laus_names$area),]

#C6-arrange columns of laus_names in their final order
laus_names %>%
  select(areaname,current_laborforce,current_emplab,current_unemp,current_unemprate
         ,previous_laborforce,previous_emplab,previous_unemp,previous_unemprate
  ) -> laus_table

#C7-format rates with one decimal point and counts with commas
laus_table[,c(2:4,6:8)]<-format(laus_table[,c(2:4,6:8)],big.mark = ',', trim = TRUE)
laus_table[,c(5,9)]<-round(laus_table[,c(5,9)],1)



#--------------------------------------------------WORD FORMATTING------------------------------------------------------------------


#create word doc object
ls<-docx(template = 'laborshed_template2.docx')

#add initial page break
ls<-addPageBreak(ls)

#define laborshed doc properties
qf_header_text<-textProperties(color = 'black'            #
                               ,font.size = 7           #
                               ,font.family = 'Arial'     #
                               ,font.weight = 'bold')     #

qf_body_text<-textProperties(color = 'black'                 #
                             ,font.size = 6.5                 #
                             ,font.family = 'Arial')          #

body_text<-textProperties(color = 'black'                 #
                          ,font.size = 10.5               #
                          ,font.family = 'Arial')         #

par_body_text<-textProperties(color = 'black'                 #
                          ,font.size = 12               #
                          ,font.family = 'Arial')


body_header_text<-textProperties(color = 'black'          #
                          ,font.size = 10.5               #
                          ,font.family = 'Arial'
                          ,font.weight = 'bold')          #

title_format <-textProperties(color = 'black'             #
                  ,font.size = 14                       #
                  ,font.family = 'Arial'
                   ,font.weight = 'bold')  

lp_header_text<-textProperties(color = 'black'                #
                            ,font.size = 6.5            #
                            ,font.family = 'Arial'      #
                            ,font.weight = 'bold')      #<-----------Text Properties
lp_body_text<-textProperties(color = 'black'                  #
                          ,font.size = 6.5                 #
                          ,font.family = 'Arial')          #

subheader_text<-textProperties(color = 'black'            #
                               ,font.size = 6.5           #
                               ,font.family = 'Arial'     #
                               ,font.weight = 'bold'      #<-------Text Properites
                               ,font.style = 'italic')    #
legend_text<-textProperties(color = 'black'               #
                            ,font.size = 6                #
                            ,font.family = 'Arial')       #
legend_head_text<-textProperties(color = 'black'          #
                                 ,font.size = 6           #
                                 ,font.family = 'Arial'   #
                                 ,font.weight = 'bold')   #
legend_subhead_text<-textProperties(color = 'black'       #
                                  ,font.size = 6          #
                                  ,font.family = 'Arial'  #
                                  ,font.style = 'italic') #

source_text<-textProperties(color = 'black'               #
                                    ,font.size = 7.5        #
                                    ,font.family = 'Arial'#
                                    ,font.style = 'normal')#


bold_intro_text<-textProperties(color = 'black'                 #
                                ,font.size = 16               #
                                ,font.family = 'Arial'
                                ,font.weight='bold')

intro_text<-textProperties(color = 'black'                 #
                           ,font.size = 11               #
                           ,font.family = 'Arial')



no_border<-borderProperties(style = 'none')                  #
plain_border<-borderProperties(color = 'black'               #<----------Border Properties
                               ,style = 'solid'              #
                               ,width = 1)                   #

center_align <-parProperties(text.align = 'center')     #<---------Align Text

left_align <-parProperties(text.align = 'left')

right_align <-parProperties(text.align = 'right')


area_title <- pot(value=area_name, title_format)

#--------------------------------------------------------------------------------------------------------
#                                    DATE
#--------------------------------------------------------------------------------------------------------

date <- pot(report_date,textProperties(font.family = 'Arial', font.weight = 'bold', font.size = 12))

addParagraph(ls, date, par.properties=parProperties(text.align = 'center'), bookmark = 'Date')

#--------------------------------------------------------------------------------------------------------
#                                     EXECUTIVE SUMMARY
#--------------------------------------------------------------------------------------------------------

#add initial page break


ls <- addTitle(ls, 'Executive Summary') #Executive summary title



par2 = paste('The analysis of workforce and demographic characteristics, including commuting patterns of'   #Intro
             ,area_name, 'was conducted to provide economic data on the population and labor force'
             ,'living or working in the county.  The report is useful for detailing where workers work and'
             , 'live in order to align resources.  This report includes population, labor force, and'
             , 'demographics for', paste0(area_name, '.'), sep=" ")

par2 <- pot(par2, intro_text)


par3 = paste('Workers employed in', area_name, 'are clustered in the'       #Top industries info
             , arrange(wac_table[which(wac_table$type=='ind'),], desc(share))$title[1]
             , paste0('(',sprintf("%.1f",arrange(wac_table[which(wac_table$type=='ind'),], desc(share))$share[1])
                      ,'%), and')
             , arrange(wac_table[which(wac_table$type=='ind'),], desc(share))$title[2]
             , paste0('(',sprintf("%.1f",arrange(wac_table[which(wac_table$type=='ind'),], desc(share))$share[2])
                      ,'%) industries.')
             , 'Workers living in', area_name, 'are concentrated in the'
             , arrange(rac_table[which(rac_table$type=='ind'),], desc(share))$title[1]
             , paste0('(',sprintf("%.1f",arrange(rac_table[which(rac_table$type=='ind'),], desc(share))$share[1])
                      ,'%), and')
             , arrange(rac_table[which(rac_table$type=='ind'),], desc(share))$title[2]
             , paste0('(',sprintf("%.1f",arrange(rac_table[which(rac_table$type=='ind'),], desc(share))$share[2])
                      ,'%) industries.')
             ,sep = " "
)

par3 <- pot(par3, intro_text)


par4 = paste('A detailed examination of commuting patterns for', area_name, 'shows that the county'     #Inflow stats
             ,'has a net'
             , ifelse(in_out_summ$count[in_out_summ$Title=='Net Job Inflow (+) or Outflow (-)']>0, 'inflow of','outflow of')
             , in_out_summ$count[in_out_summ$Title=='Net Job Inflow (+) or Outflow (-)'], 'workers.'
             , 'Using the latest annual Census data available, there were'
             , in_out_summ$count[1], 'workers employed in', area_name, 'and'
             , in_out_summ$count[2], 'workers living in', paste0(area_name,'.'), 'Of the workers who'
             , 'lived in the county,', in_out_summ$count[in_out_summ$Title=='Living in the Selection Area but Employed Outside']
             , 'workers', paste0('(', in_out_summ$Share[in_out_summ$Title=='Living in the Selection Area but Employed Outside'],
                                 ') were employed outside the county.')
             , sep=" ")

par4 <- pot(par4, intro_text)


par5 = paste('With', in_out_summ$Share[in_out_summ$Title=='Living in the Selection Area but Employed Outside']   #Outflow rankings
             , 'of workers who reside in', area_name, 'employed outside the county,'
             , area_name, 'had the ... [ENTER OUTFLOW RANKINGS]...', 'among Florida counties.'
             , paste(outflow_rank[1, "LaborOutflow"], 'County', sep=" ")
             , paste0('(', sprintf("%.1f",outflow_rank[1, 2]), '%)') 
             ,'had the highest outflow rate, followed by'
             , paste(outflow_rank[2, "LaborOutflow"], 'County', sep=" ")
             , paste0('(', sprintf("%.1f",outflow_rank[2, 2]), '%),')
             , 'and'
             , paste(outflow_rank[3, "LaborOutflow"], 'County', sep=" ")
             , paste0('(', sprintf("%.1f",outflow_rank[3, 2]), '%).')
             , paste(outflow_rank[nrow(outflow_rank), "LaborOutflow"], 'County', sep=" ")
             , paste0('(', sprintf("%.1f",outflow_rank[nrow(outflow_rank), 2]), '%),')
             , paste(outflow_rank[nrow(outflow_rank)-1, "LaborOutflow"], 'County', sep=" ")
             , paste0('(', sprintf("%.1f",outflow_rank[nrow(outflow_rank)-1, 2]), '%),') 
             , 'and'
             , paste(outflow_rank[nrow(outflow_rank)-3, "LaborOutflow"], 'County', sep=" ")
             , paste0('(', sprintf("%.1f",outflow_rank[nrow(outflow_rank)-3, 2]), '%)') 
             ,'had the lowest worker outflow rates.'
             , sep = " ")

par5 <- pot(par5, intro_text)


par6 = paste('Of the', in_out_summ$count[in_out_summ$Title=='Living in the Selection Area but Employed Outside']    #Work and home destinations
             , area_name, 'workers employed outside the county, the top destination counties are'
             , sub('\\s*,.*','', arrange(work_dest[which(work_dest$title!=paste0(area_name,', FL') & work_dest$title!='All Other Locations'
                                                               & work_dest$title!='Total Primary Jobs'),], desc(count))$title[1])
             , paste0('(', comma(arrange(work_dest[which(work_dest$title!=paste0(area_name,', FL') & work_dest$title!='All Other Locations'
                                                         & work_dest$title!='Total Primary Jobs'),], desc(count))$count[1]),' workers),')
             , sub('\\s*,.*','', arrange(work_dest[which(work_dest$title!=paste0(area_name,', FL') & work_dest$title!='All Other Locations'
                                                         & work_dest$title!='Total Primary Jobs'),], desc(count))$title[2])
             , paste0('(', comma(arrange(work_dest[which(work_dest$title!=paste0(area_name,', FL') & work_dest$title!='All Other Locations'
                                                         & work_dest$title!='Total Primary Jobs'),], desc(count))$count[2]),' workers),')
             , 'and'
             , sub('\\s*,.*','', arrange(work_dest[which(work_dest$title!=paste0(area_name,', FL') & work_dest$title!='All Other Locations'
                                                         & work_dest$title!='Total Primary Jobs'),], desc(count))$title[3])
             , paste0('(', comma(arrange(work_dest[which(work_dest$title!=paste0(area_name,', FL') & work_dest$title!='All Other Locations'
                                                         & work_dest$title!='Total Primary Jobs'),], desc(count))$count[3]),' workers).')
             , 'Of the', in_out_summ$count[in_out_summ$Title=='Employed in the Selection Area but Living Outside']
             , area_name, 'workers living outside the county, the top origin counties are'
             , sub('\\s*,.*','', arrange(home_dest[which(home_dest$title!=paste0(area_name,', FL') & home_dest$title!='All Other Locations'
                                                         & home_dest$title!='Total Primary Jobs'),], desc(count))$title[1])
             , paste0('(', comma(arrange(home_dest[which(home_dest$title!=paste0(area_name,', FL') & home_dest$title!='All Other Locations'
                                                         & home_dest$title!='Total Primary Jobs'),], desc(count))$count[1]),' workers),')
             , sub('\\s*,.*','', arrange(home_dest[which(home_dest$title!=paste0(area_name,', FL') & home_dest$title!='All Other Locations'
                                                         & home_dest$title!='Total Primary Jobs'),], desc(count))$title[2])
             , paste0('(', comma(arrange(home_dest[which(home_dest$title!=paste0(area_name,', FL') & home_dest$title!='All Other Locations'
                                                         & home_dest$title!='Total Primary Jobs'),], desc(count))$count[2]),' workers),')
             , 'and'
             , sub('\\s*,.*','', arrange(home_dest[which(home_dest$title!=paste0(area_name,', FL') & home_dest$title!='All Other Locations'
                                                         & home_dest$title!='Total Primary Jobs'),], desc(count))$title[3])
             , paste0('(', comma(arrange(home_dest[which(home_dest$title!=paste0(area_name,', FL') & home_dest$title!='All Other Locations'
                                                         & home_dest$title!='Total Primary Jobs'),], desc(count))$count[3]),' workers).')
)


par6 <- pot(par6, intro_text)     


exec_summ_pars = set_of_paragraphs(par2,par3,par4,par5,par6)

addParagraph(ls, "")
addParagraph(ls, exec_summ_pars, par.properties = parProperties(text.align = 'justify', padding.bottom = 12))

#--------------------------------------------------------------------------------------------------------
#                                      QUICKFACTS TABLE
#--------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)

#create header flex table----------------------------------------------------
f_header_table<-FlexTable(qf_header
                          ,body.text.props = qf_header_text
                          ,header.columns = FALSE)
#set borders
f_header_table<-setFlexTableBorders(f_header_table
                                    ,inner.vertical = no_border
                                    ,inner.horizontal = no_border
                                    ,outer.vertical = plain_border
                                    ,outer.horizontal = plain_border)
#set column widths
f_header_table<-setFlexTableWidths(f_header_table,c(1,3.925,(1.45/2),.85))
#align text
f_header_table[1,1:4]<-right_align

#add to doc object
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'People QuickFacts', level=1)
ls<-addFlexTable(ls,f_header_table,par.properties = center_align)

#create first page flex table -------------------------------------------------------------
#combine population, age, race, and characteristic tables
page1<-rbind.data.frame(qf_pop_table,age_table,race_table,char_table)

#insert into flex table
f_page1<-FlexTable(page1
                   ,body.text.props = qf_body_text
                   ,header.columns = FALSE)
#set borders
f_page1<-setFlexTableBorders(f_page1
                             ,inner.vertical = no_border
                             ,inner.horizontal = no_border
                             ,outer.vertical = plain_border
                             ,outer.horizontal = plain_border)
#set widths
f_page1<-setFlexTableWidths(f_page1,c(4.2,(1.45/2),(1.45/2),.85))
#set subheader format
f_page1[c(1,9,14,23),]<-subheader_text
#align text
f_page1[,2:4]<-right_align
#add to doc object
ls<-addFlexTable(ls,f_page1,par.properties = center_align)

#create legend flex table---------------------------------------------------------------------
f_legend_table<-FlexTable(legend
                          ,body.text.props = legend_text
                          ,header.columns = FALSE)
#set widths
f_legend_table<-setFlexTableWidths(f_legend_table,6.5)
#set borders
f_legend_table<-setFlexTableBorders(f_legend_table
                                    ,inner.vertical = no_border
                                    ,inner.horizontal = no_border
                                    ,outer.vertical = plain_border
                                    ,outer.horizontal = plain_border)
#set subheader format
f_legend_table[1,]<-legend_head_text
f_legend_table[c(3,7),]<-legend_subhead_text
#add source footer
f_legend_table<- addFooterRow(f_legend_table, value = "Source:  U.S. Census Bureau, State & County QuickFacts."
                             ,colspan=1
                             ,text.properties=source_text)
f_legend_table[to = 'footer', side = 'right'] <- borderProperties(style='none')
f_legend_table[to = 'footer', side = 'left'] <- borderProperties(style='none')
f_legend_table[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
f_legend_table[to = 'footer', side = 'top'] <- borderProperties(style='solid')
#add to doc object
ls<-addFlexTable(ls,f_legend_table,par.properties = parProperties(text.align = 'center'))


#PAGE 2----------------------------------------------------------------------------------------
ls<-addPageBreak(ls)

#header table
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addParagraph(ls, value = pot('People QuickFacts (continued)', title_format), stylename = 'Normal', 
                   par.properties = parProperties(text.align='left', padding.top = 8))
ls<-addFlexTable(ls,f_header_table,par.properties = parProperties(text.align = 'center'))

#page 2 table
page2<-rbind.data.frame(hous_table,fam_table,edu_table,heal_table,econ_table,tran_table,inc_table)

#insert into flex table
f_page2<-FlexTable(page2
                   ,body.text.props = qf_body_text
                   ,header.columns = FALSE)
#set borders
f_page2<-setFlexTableBorders(f_page2
                             ,inner.vertical = no_border
                             ,inner.horizontal = no_border
                             ,outer.vertical = plain_border
                             ,outer.horizontal = plain_border)
#set widths
f_page2<-setFlexTableWidths(f_page2,c(4.2,(1.45/2),(1.45/2),.85))
#set subheader format
f_page2[c(1,9,14,17,20,29,31),]<-subheader_text
#align text
f_page2[,2:4]<-right_align

#add to doc object
ls<-addFlexTable(ls,f_page2,par.properties = center_align)

#add legend to page 2 of doc object
ls<-addFlexTable(ls,f_legend_table,par.properties = parProperties(text.align = 'center'))

#PAGE 3-----------------------------------------------------------------------------------
ls<-addPageBreak(ls)

#header table
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'Business and Geography QuickFacts', level=1)
ls<-addFlexTable(ls,f_header_table,par.properties = parProperties(text.align = 'center'))

#page 3 table
page3<-rbind.data.frame(bus_table,geo_table)

#insert into flex table
f_page3<-FlexTable(page3
                   ,body.text.props = qf_body_text
                   ,header.columns = FALSE)
#set borders
f_page3<-setFlexTableBorders(f_page3
                             ,inner.vertical = no_border
                             ,inner.horizontal = no_border
                             ,outer.vertical = plain_border
                             ,outer.horizontal = plain_border)
#set widths
f_page3<-setFlexTableWidths(f_page3,c(4.2,(1.45/2),(1.45/2),.85))
#set subheader format
f_page3[c(1,14),]<-subheader_text
#align text
f_page3[,2:4]<-right_align
#add source footer
f_page3<- addFooterRow(f_page3, value = "Source:  U.S. Census Bureau, State & County QuickFacts."
                              ,colspan=4
                              ,text.properties=source_text)
f_page3[to = 'footer', side = 'right'] <- borderProperties(style='none')
f_page3[to = 'footer', side = 'left'] <- borderProperties(style='none')
f_page3[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
f_page3[to = 'footer', side = 'top'] <- borderProperties(style='solid')
#add to doc object
ls<-addFlexTable(ls,f_page3,par.properties = center_align)

#add to legend to page 3 of doc object
#ls<-addFlexTable(ls,f_legend_table,par.properties = parProperties(text.align = 'center'))

#--------------------------------------------------------------------------------------------------------
#                                      LAUS AND POPULATION TABLES
#--------------------------------------------------------------------------------------------------------

#E3-laus flextable
#E3a-add data to table
ft_laus<-FlexTable(laus_table,header.columns = FALSE,
                   body.text.props = lp_body_text, body.par.props = left_align)

#E3b-add outside borders
ft_laus<-setFlexTableBorders(ft_laus
                             ,inner.vertical = no_border
                             ,inner.horizontal = no_border
                             ,outer.vertical = plain_border
                             ,outer.horizontal = plain_border)

#E3c-set column widths
ft_laus<-setFlexTableWidths(ft_laus,c(1.5,.65,.65,.6,.6,.65,.65,.6,.6))

#E3d-align text 
ft_laus[,c(2:9)]<-right_align

#E3e-create header rows
laus_head1<-FlexRow(c('','FORCE','MENT','LEVEL','RATE (%)','FORCE','MENT','LEVEL','RATE (%)')
                    ,text.properties = lp_header_text)
laus_head2<-FlexRow(c('','LABOR','EMPLOY-','UNEMPLOYMENT','LABOR','EMPLOY-','UNEMPLOYMENT')
                    ,colspan = c(1,1,1,2,1,1,2)
                    ,text.properties = lp_header_text)
laus_head3<-FlexRow(c('',month_year,month_prev_year)
                    ,colspan = c(1,4,4)
                    ,text.properties = lp_header_text)

#E3f-add header rows
ft_laus<-addHeaderRow(ft_laus,laus_head3)
ft_laus<-addHeaderRow(ft_laus,laus_head2)
ft_laus<-addHeaderRow(ft_laus,laus_head1)

#E3g-header borders
ft_laus[c(1,2),1, to = 'header', side = 'bottom']<-no_border
ft_laus[2,,to = 'header', side = 'bottom']<-no_border
ft_laus[2,c(2,3,6,7),to = 'header', side = 'right']<-no_border
ft_laus[3,c(2:4,6:8),to = 'header', side = 'right']<-no_border
#E3h-header alignment
ft_laus[1:3,,to = 'header']<-center_align

#E3h-remaining body borders
ft_laus[,1, side = 'right']<-plain_border
ft_laus[,5, side = 'right']<-plain_border

#E3i-color county row of interest
color_row<-which(laus_table$areaname == area_name)
shade<-rgb(red = 195,green = 224,blue = 242,maxColorValue = 255)
ft_laus<-setFlexTableBackgroundColors(ft_laus,i = color_row,colors = shade)

#add source footer
ft_laus<- addFooterRow(ft_laus, value = "Source:  U.S. Department of Labor, Bureau of Labor Statistics, Local Area Unemployment Statistics."
                       ,colspan=9
                       ,text.properties=source_text)
ft_laus[to = 'footer', side = 'right'] <- borderProperties(style='none')
ft_laus[to = 'footer', side = 'left'] <- borderProperties(style='none')
ft_laus[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
ft_laus[to = 'footer', side = 'top'] <- borderProperties(style='solid')

#E4-population flex table
#E4a-add data to table
ft_pop<-FlexTable(pop_table,header.columns = FALSE,
                  body.text.props = lp_body_text)

#E4b-add outside borders
ft_pop<-setFlexTableBorders(ft_pop
                            ,inner.vertical = no_border
                            ,inner.horizontal = no_border
                            ,outer.vertical = plain_border
                            ,outer.horizontal = plain_border)

#E4c-set column widths
ft_pop<-setFlexTableWidths(ft_pop,c(1.5,(5/7),(5/7),(5/7),(5/7),(5/7),(5/7),(5/7)))

#E4d-align text 
ft_pop[,c(2:8)]<-right_align

#E4e-create header rows

years <- pop_data_range
                             
pop_head1<-FlexRow(c('','Population Estimate (as of July 1)')
                   ,colspan = c(1,7)
                   ,text.properties = lp_header_text)
pop_head2<-FlexRow(c('',years)
                   ,text.properties = lp_header_text)

#E4f-add header rows
ft_pop<-addHeaderRow(ft_pop,pop_head1)
ft_pop<-addHeaderRow(ft_pop,pop_head2)

#E4g-header borders
ft_pop[1,1,to = 'header',side = 'bottom']<-no_border
ft_pop[2,2:7,to = 'header',side = 'right']<-no_border

#E4h-header alignment
ft_pop[1:2,,to = 'header']<-center_align

#E4h-remaining body borders
ft_pop[,1, side = 'right']<-plain_border

#E4i-color county row of interest
color_row<-which(pop_table$areaname == area_name)
ft_pop<-setFlexTableBackgroundColors(ft_pop,i = color_row,colors = shade)

#add source footer
ft_pop<- addFooterRow(ft_pop, value = "Source:  U.S. Census Bureau, Population Division."
                       ,colspan=8
                       ,text.properties=source_text)
ft_pop[to = 'footer', side = 'right'] <- borderProperties(style='none')
ft_pop[to = 'footer', side = 'left'] <- borderProperties(style='none')
ft_pop[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
ft_pop[to = 'footer', side = 'top'] <- borderProperties(style='solid')

ls <- addParagraph(ls, value = "", par.properties = parProperties(padding.bottom = 8))
ls <- addTitle(ls, 'Labor Force, Employment, and Unemployment', level=1)
ls <- addFlexTable(ls,ft_laus,par.properties = center_align)
ls <- addParagraph(ls, value = "", par.properties = parProperties(padding.bottom = 8))
ls <- addTitle(ls, 'Population', level=1)
ls<-addFlexTable(ls,ft_pop,par.properties = center_align)

#---------------------------------------------------------------------------------------------------------
#                                     PAGE FOR MAP ABOUT WHERE WORKERS WORK
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)

ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, paste('Where Workers Work in', area_name, sep=" "), level=1)

ls <- addParagraph(ls, pot("Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.",source_text),
                         par.properties = parProperties(padding.top = 12))

#---------------------------------------------------------------------------------------------------------
#                                      WORK AREA PROFILE
#---------------------------------------------------------------------------------------------------------
  
wac_table1 <- FlexTable(wac_table_final[1,], body.text.props = body_text, header.columns = FALSE)
wac_table2 <- FlexTable(wac_table_final[2:4,], body.text.props = body_text, header.columns = FALSE)
wac_table3 <- FlexTable(wac_table_final[5:7,], body.text.props = body_text, header.columns = FALSE)
wac_table4 <- FlexTable(wac_table_final[8:27,], body.text.props = body_text, header.columns = FALSE)


header_row = FlexRow()
header_row2 = FlexRow()
header_row3 = FlexRow()
header_row4 = FlexRow()


header_row[1] = FlexCell("")

header_row[2] = FlexCell(pot("Count", body_header_text),
                         par.properties = parProperties(text.align='right'))

header_row[3] = FlexCell(pot("Share", body_header_text),
                         par.properties = parProperties(text.align='right'))



header_row2[1] = FlexCell(pot("Jobs by Worker Age", body_header_text))
header_row2[2] = FlexCell("")
header_row2[3] = FlexCell("")


header_row3[1] = FlexCell(pot("Jobs by Earnings", body_header_text))
header_row3[2] = FlexCell("")
header_row3[3] = FlexCell("")


header_row4[1] = FlexCell(pot("Jobs by NAICS Industry Sector", body_header_text))
header_row4[2] = FlexCell("")
header_row4[3] = FlexCell("")



wac_table1 <- addHeaderRow(wac_table1, header_row)
wac_table1 <- addFooterRow(wac_table1, value = "", colspan=3)
wac_table2 <- addHeaderRow(wac_table2, header_row2)
wac_table2 <- addFooterRow(wac_table2, value = "", colspan=3)
wac_table3 <- addHeaderRow(wac_table3, header_row3)
wac_table3 <- addFooterRow(wac_table3, value = "", colspan=3)
wac_table4 <- addHeaderRow(wac_table4, header_row4)
wac_table4 <- addFooterRow(wac_table4, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                           text.properties=source_text)


wac_table1 <- setFlexTableBorders(wac_table1
                                     ,inner.vertical = no_border
                                     ,inner.horizontal = no_border
                                     ,outer.vertical = no_border
                                    ,outer.horizontal = no_border)

wac_table2 <- setFlexTableBorders(wac_table2
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

wac_table3 <- setFlexTableBorders(wac_table3
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

wac_table4 <- setFlexTableBorders(wac_table4
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)


wac_table1 <- setFlexTableWidths(wac_table1,c(4.5,1,1))
wac_table2 <- setFlexTableWidths(wac_table2,c(4.5,1,1))
wac_table3 <- setFlexTableWidths(wac_table3,c(4.5,1,1))
wac_table4 <- setFlexTableWidths(wac_table4,c(4.5,1,1))


wac_table1[1,1] = body_header_text

wac_table1[,1,side="left"] <- borderProperties(style='solid')
wac_table1[,3,side="right"] <- borderProperties(style='solid')
wac_table1[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table1[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table1[1,, to='header', side="top"] <- borderProperties(style='solid')
wac_table1[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table1[to = 'footer', side = 'top'] <- borderProperties(style='none')


wac_table2[,1,side="left"] <- borderProperties(style='solid')
wac_table2[,3,side="right"] <- borderProperties(style='solid')
wac_table2[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table2[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table2[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table2[to = 'footer', side = 'top'] <- borderProperties(style='none')


wac_table3[,1,side="left"] <- borderProperties(style='solid')
wac_table3[,3,side="right"] <- borderProperties(style='solid')
wac_table3[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table3[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table3[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table3[to = 'footer', side = 'top'] <- borderProperties(style='none')


wac_table4[,1,side="left"] <- borderProperties(style='solid')
wac_table4[,3,side="right"] <- borderProperties(style='solid')
wac_table4[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table4[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table4[to = 'footer', side = 'right'] <- borderProperties(style='none')
wac_table4[to = 'footer', side = 'left'] <- borderProperties(style='none')
wac_table4[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table4[to = 'footer', side = 'top'] <- borderProperties(style='solid')




wac_table1[, 1] <- left_align
wac_table1[,2:3] <- right_align

wac_table2[, 1] <- left_align
wac_table2[,2:3] <- right_align

wac_table3[, 1] <- left_align
wac_table3[,2:3] <- right_align

wac_table4[, 1] <- left_align
wac_table4[,2:3] <- right_align


ls <- addPageBreak(ls)
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'Work Area Profile Summary', level=1)

ls <- addFlexTable(ls, wac_table1)
ls <- addFlexTable(ls, wac_table2)
ls <- addFlexTable(ls, wac_table3)
ls <- addFlexTable(ls, wac_table4)

#-----SECOND PAGE OF WORK AREA PROFILE 

ls <- addPageBreak(ls)
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls,'Work Area Profile Summary - Demographics', level=1)

wac_table5 <- FlexTable(wac_table_final[28:33,], body.text.props = body_text, header.columns = FALSE)
wac_table6 <- FlexTable(wac_table_final[34:35,], body.text.props = body_text, header.columns = FALSE)
wac_table7 <- FlexTable(wac_table_final[36:40,], body.text.props = body_text, header.columns = FALSE)
wac_table8 <- FlexTable(wac_table_final[41:42,], body.text.props = body_text, header.columns = FALSE)


header_row5 = FlexRow()
header_row6 = FlexRow()
header_row7 = FlexRow()
header_row8 = FlexRow()


header_row5[1] = FlexCell(pot("Jobs by Worker Race", body_header_text))
header_row5[2] = FlexCell("")
header_row5[3] = FlexCell("")


header_row6[1] = FlexCell(pot("Jobs by Worker Ethnicity", body_header_text))
header_row6[2] = FlexCell("")
header_row6[3] = FlexCell("")


header_row7[1] = FlexCell(pot("Jobs by Worker Educational Attainment", body_header_text))
header_row7[2] = FlexCell("")
header_row7[3] = FlexCell("")


header_row8[1] = FlexCell(pot("Jobs by Worker Sex", body_header_text))
header_row8[2] = FlexCell("")
header_row8[3] = FlexCell("")

wac_table5 <- addHeaderRow(wac_table5, header_row5)
wac_table5 <- addFooterRow(wac_table5, value = "", colspan=3)
wac_table6 <- addHeaderRow(wac_table6, header_row6)
wac_table6 <- addFooterRow(wac_table6, value = "", colspan=3)
wac_table7 <- addHeaderRow(wac_table7, header_row7)
wac_table7 <- addFooterRow(wac_table7, value = "", colspan=3)
wac_table8 <- addHeaderRow(wac_table8, header_row8)
wac_table8 <- addFooterRow(wac_table8, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                           text.properties=source_text)


wac_table5 <- setFlexTableBorders(wac_table5
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

wac_table6 <- setFlexTableBorders(wac_table6
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

wac_table7 <- setFlexTableBorders(wac_table7
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

wac_table8 <- setFlexTableBorders(wac_table8
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)


wac_table5 <- setFlexTableWidths(wac_table5,c(4.5,1,1))
wac_table6 <- setFlexTableWidths(wac_table6,c(4.5,1,1))
wac_table7 <- setFlexTableWidths(wac_table7,c(4.5,1,1))
wac_table8 <- setFlexTableWidths(wac_table8,c(4.5,1,1))

wac_table5[,1,side="left"] <- borderProperties(style='solid')
wac_table5[,3,side="right"] <- borderProperties(style='solid')
wac_table5[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table5[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table5[1,, to='header', side="top"] <- borderProperties(style='solid')
wac_table5[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table5[to = 'footer', side = 'top'] <- borderProperties(style='none')


wac_table6[,1,side="left"] <- borderProperties(style='solid')
wac_table6[,3,side="right"] <- borderProperties(style='solid')
wac_table6[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table6[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table6[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table6[to = 'footer', side = 'top'] <- borderProperties(style='none')


wac_table7[,1,side="left"] <- borderProperties(style='solid')
wac_table7[,3,side="right"] <- borderProperties(style='solid')
wac_table7[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table7[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table7[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table7[to = 'footer', side = 'top'] <- borderProperties(style='none')


wac_table8[,1,side="left"] <- borderProperties(style='solid')
wac_table8[,3,side="right"] <- borderProperties(style='solid')
wac_table8[1,1,to='header', side="left"] <- borderProperties(style='solid')
wac_table8[1,3,to='header', side="right"] <- borderProperties(style='solid')
wac_table8[to = 'footer', side = 'right'] <- borderProperties(style='none')
wac_table8[to = 'footer', side = 'left'] <- borderProperties(style='none')
wac_table8[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wac_table8[to = 'footer', side = 'top'] <- borderProperties(style='solid')


wac_table5[, 1] <- left_align
wac_table5[,2:3] <- right_align

wac_table6[, 1] <- left_align
wac_table6[,2:3] <- right_align

wac_table7[, 1] <- left_align
wac_table7[,2:3] <- right_align

wac_table8[, 1] <- left_align
wac_table8[,2:3] <- right_align


ls <- addFlexTable(ls, wac_table5)
ls <- addFlexTable(ls, wac_table6)
ls <- addFlexTable(ls, wac_table7)
ls <- addFlexTable(ls, wac_table8)


#---------------------------------------------------------------------------------------------------------
#                                     PAGE FOR MAP ABOUT WHERE WORKERS LIVE
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)

ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, paste('Where Workers Live in', area_name, sep=" "), level=1)

ls <- addParagraph(ls, pot("Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.",source_text),
                   par.properties = parProperties(padding.top = 12))


#---------------------------------------------------------------------------------------------------------
#                                      HOME AREA PROFILE
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'Home Area Profile Summary', level=1)


rac_table1 <- FlexTable(rac_table_final[1,], body.text.props = body_text, header.columns = FALSE)
rac_table2 <- FlexTable(rac_table_final[2:4,], body.text.props = body_text, header.columns = FALSE)
rac_table3 <- FlexTable(rac_table_final[5:7,], body.text.props = body_text, header.columns = FALSE)
rac_table4 <- FlexTable(rac_table_final[8:27,], body.text.props = body_text, header.columns = FALSE)


header_row = FlexRow()
header_row2 = FlexRow()
header_row3 = FlexRow()
header_row4 = FlexRow()


header_row[1] = FlexCell("")

header_row[2] = FlexCell(pot("Count", body_header_text),
                         par.properties = parProperties(text.align='right'))

header_row[3] = FlexCell(pot("Share", textProperties(font.weight = 'bold')),
                         par.properties = parProperties(text.align='right'))



header_row2[1] = FlexCell(pot("Jobs by Worker Age", body_header_text))
header_row2[2] = FlexCell("")
header_row2[3] = FlexCell("")


header_row3[1] = FlexCell(pot("Jobs by Earnings", body_header_text))
header_row3[2] = FlexCell("")
header_row3[3] = FlexCell("")


header_row4[1] = FlexCell(pot("Jobs by NAICS Industry Sector", body_header_text))
header_row4[2] = FlexCell("")
header_row4[3] = FlexCell("")



rac_table1 <- addHeaderRow(rac_table1, header_row)
rac_table1 <- addFooterRow(rac_table1, value = "", colspan=3)
rac_table2 <- addHeaderRow(rac_table2, header_row2)
rac_table2 <- addFooterRow(rac_table2, value = "", colspan=3)
rac_table3 <- addHeaderRow(rac_table3, header_row3)
rac_table3 <- addFooterRow(rac_table3, value = "", colspan=3)
rac_table4 <- addHeaderRow(rac_table4, header_row4)
rac_table4 <- addFooterRow(rac_table4, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                           text.properties=source_text)


rac_table1 <- setFlexTableBorders(rac_table1
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

rac_table2 <- setFlexTableBorders(rac_table2
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

rac_table3 <- setFlexTableBorders(rac_table3
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

rac_table4 <- setFlexTableBorders(rac_table4
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)


rac_table1 <- setFlexTableWidths(rac_table1,c(4.5,1,1))
rac_table2 <- setFlexTableWidths(rac_table2,c(4.5,1,1))
rac_table3 <- setFlexTableWidths(rac_table3,c(4.5,1,1))
rac_table4 <- setFlexTableWidths(rac_table4,c(4.5,1,1))


rac_table1[1,1] = body_header_text

rac_table1[,1,side="left"] <- borderProperties(style='solid')
rac_table1[,3,side="right"] <- borderProperties(style='solid')
rac_table1[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table1[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table1[1,, to='header', side="top"] <- borderProperties(style='solid')
rac_table1[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table1[to = 'footer', side = 'top'] <- borderProperties(style='none')


rac_table2[,1,side="left"] <- borderProperties(style='solid')
rac_table2[,3,side="right"] <- borderProperties(style='solid')
rac_table2[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table2[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table2[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table2[to = 'footer', side = 'top'] <- borderProperties(style='none')


rac_table3[,1,side="left"] <- borderProperties(style='solid')
rac_table3[,3,side="right"] <- borderProperties(style='solid')
rac_table3[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table3[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table3[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table3[to = 'footer', side = 'top'] <- borderProperties(style='none')


rac_table4[,1,side="left"] <- borderProperties(style='solid')
rac_table4[,3,side="right"] <- borderProperties(style='solid')
rac_table4[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table4[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table4[to = 'footer', side = 'right'] <- borderProperties(style='none')
rac_table4[to = 'footer', side = 'left'] <- borderProperties(style='none')
rac_table4[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table4[to = 'footer', side = 'top'] <- borderProperties(style='solid')




rac_table1[, 1] <- left_align
rac_table1[,2:3] <- right_align

rac_table2[, 1] <- left_align
rac_table2[,2:3] <- right_align

rac_table3[, 1] <- left_align
rac_table3[,2:3] <- right_align

rac_table4[, 1] <- left_align
rac_table4[,2:3] <- right_align


ls <- addFlexTable(ls, rac_table1)
ls <- addFlexTable(ls, rac_table2)
ls <- addFlexTable(ls, rac_table3)
ls <- addFlexTable(ls, rac_table4)

#-----SECOND PAGE OF HOME AREA PROFILE 

ls <- addPageBreak(ls)
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'Home Area Profile Summary - Demographics', level=1)

rac_table5 <- FlexTable(rac_table_final[28:33,], body.text.props = body_text, header.columns = FALSE)
rac_table6 <- FlexTable(rac_table_final[34:35,], body.text.props = body_text, header.columns = FALSE)
rac_table7 <- FlexTable(rac_table_final[36:40,], body.text.props = body_text, header.columns = FALSE)
rac_table8 <- FlexTable(rac_table_final[41:42,], body.text.props = body_text, header.columns = FALSE)


header_row5 = FlexRow()
header_row6 = FlexRow()
header_row7 = FlexRow()
header_row8 = FlexRow()


header_row5[1] = FlexCell(pot("Jobs by Worker Race", body_header_text))
header_row5[2] = FlexCell("")
header_row5[3] = FlexCell("")


header_row6[1] = FlexCell(pot("Jobs by Worker Ethnicity", body_header_text))
header_row6[2] = FlexCell("")
header_row6[3] = FlexCell("")


header_row7[1] = FlexCell(pot("Jobs by Worker Educational Attainment", body_header_text))
header_row7[2] = FlexCell("")
header_row7[3] = FlexCell("")


header_row8[1] = FlexCell(pot("Jobs by Worker Sex", body_header_text))
header_row8[2] = FlexCell("")
header_row8[3] = FlexCell("")

rac_table5 <- addHeaderRow(rac_table5, header_row5)
rac_table5 <- addFooterRow(rac_table5, value = "", colspan=3)
rac_table6 <- addHeaderRow(rac_table6, header_row6)
rac_table6 <- addFooterRow(rac_table6, value = "", colspan=3)
rac_table7 <- addHeaderRow(rac_table7, header_row7)
rac_table7 <- addFooterRow(rac_table7, value = "", colspan=3)
rac_table8 <- addHeaderRow(rac_table8, header_row8)
rac_table8 <- addFooterRow(rac_table8, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                           text.properties=source_text)


rac_table5 <- setFlexTableBorders(rac_table5
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

rac_table6 <- setFlexTableBorders(rac_table6
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

rac_table7 <- setFlexTableBorders(rac_table7
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

rac_table8 <- setFlexTableBorders(rac_table8
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)


rac_table5 <- setFlexTableWidths(rac_table5,c(4.5,1,1))
rac_table6 <- setFlexTableWidths(rac_table6,c(4.5,1,1))
rac_table7 <- setFlexTableWidths(rac_table7,c(4.5,1,1))
rac_table8 <- setFlexTableWidths(rac_table8,c(4.5,1,1))

rac_table5[,1,side="left"] <- borderProperties(style='solid')
rac_table5[,3,side="right"] <- borderProperties(style='solid')
rac_table5[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table5[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table5[1,, to='header', side="top"] <- borderProperties(style='solid')
rac_table5[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table5[to = 'footer', side = 'top'] <- borderProperties(style='none')


rac_table6[,1,side="left"] <- borderProperties(style='solid')
rac_table6[,3,side="right"] <- borderProperties(style='solid')
rac_table6[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table6[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table6[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table6[to = 'footer', side = 'top'] <- borderProperties(style='none')


rac_table7[,1,side="left"] <- borderProperties(style='solid')
rac_table7[,3,side="right"] <- borderProperties(style='solid')
rac_table7[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table7[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table7[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table7[to = 'footer', side = 'top'] <- borderProperties(style='none')


rac_table8[,1,side="left"] <- borderProperties(style='solid')
rac_table8[,3,side="right"] <- borderProperties(style='solid')
rac_table8[1,1,to='header', side="left"] <- borderProperties(style='solid')
rac_table8[1,3,to='header', side="right"] <- borderProperties(style='solid')
rac_table8[to = 'footer', side = 'right'] <- borderProperties(style='none')
rac_table8[to = 'footer', side = 'left'] <- borderProperties(style='none')
rac_table8[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
rac_table8[to = 'footer', side = 'top'] <- borderProperties(style='solid')


rac_table5[, 1] <- left_align
rac_table5[,2:3] <- right_align

rac_table6[, 1] <- left_align
rac_table6[,2:3] <- right_align

rac_table7[, 1] <- left_align
rac_table7[,2:3] <- right_align

rac_table8[, 1] <- left_align
rac_table8[,2:3] <- right_align


ls <- addFlexTable(ls, rac_table5)
ls <- addFlexTable(ls, rac_table6)
ls <- addFlexTable(ls, rac_table7)
lS <- addFlexTable(ls, rac_table8)

#---------------------------------------------------------------------------------------------------------
#                                      PARAGRAPH ACCOMPANYING INFLOW/OUTFLOW MAP
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)

ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'Labor Inflow/Outflow Map', level=1)

para <- paste0('Of the ', in_out_summ$count[1], ' people who are employed in ',
  area_name, ', ', in_out_summ$count[in_out_summ$Title=='Employed and Living in the Selection Area'], ' (',
  in_out_summ$Share[in_out_summ$Title=='Employed and Living in the Selection Area'], ') live and work in the county. ', 
  'There are ', in_out_summ$count[in_out_summ$Title=='Employed in the Selection Area but Living Outside'], ' (',
  in_out_summ$Share[in_out_summ$Title=='Employed in the Selection Area but Living Outside'], ') workers who live outside ',
  area_name, ' and work within the county.', ' Of the ', in_out_summ$count[2], 
  ' workers who live within ', area_name, ', ', in_out_summ$count[in_out_summ$Title=='Living in the Selection Area but Employed Outside'], ' (',
  in_out_summ$Share[in_out_summ$Title=='Living in the Selection Area but Employed Outside'], ') residents work outside the county. ',
  'There is a net ', ifelse(in_out_summ$count[in_out_summ$Title=='Net Job Inflow (+) or Outflow (-)']>0, 'inflow of ','outflow of '), 
  in_out_summ$count[in_out_summ$Title=='Net Job Inflow (+) or Outflow (-)'], 
  ifelse(in_out_summ$count[in_out_summ$Title=='Net Job Inflow (+) or Outflow (-)']>0, ' workers into ',' workers from '), 
  area_name, '.')


ls <- addParagraph(ls, pot("Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.",source_text),
                   par.properties = parProperties(padding.bottom = 12, padding.top=12))

ls <- addParagraph(ls, value = pot(para, par_body_text), stylename = 'Normal')




#---------------------------------------------------------------------------------------------------------
#                                      INFLOW/OUTFLOW SUMMARY
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addTitle(ls, 'Labor Inflow/Outflow Summary', level=1)

io_table1 <- FlexTable(in_out_summ[1:3,], body.text.props = body_text, header.columns = FALSE)
io_table2 <- FlexTable(in_out_summ[4:6,], body.text.props = body_text, header.columns = FALSE)
io_table3 <- FlexTable(in_out_summ[7:9,], body.text.props = body_text, header.columns = FALSE)
io_table4 <- FlexTable(in_out_summ[10:19,], body.text.props = body_text, header.columns = FALSE)


header_row = FlexRow()
header_row2 = FlexRow()
header_row3 = FlexRow()
header_row4 = FlexRow()


header_row[1] = FlexCell(pot("Selection Area Labor Market Size (Primary Jobs)", body_header_text),
                         par.properties = parProperties(text.align='left'))

header_row[2] = FlexCell(pot("Count", body_header_text),
                         par.properties = parProperties(text.align='right'))

header_row[3] = FlexCell(pot("Share", body_header_text),
                         par.properties = parProperties(text.align='right'))



header_row2[1] = FlexCell(pot("In-Area Labor Force Efficiency (Primary Jobs)", body_header_text))
header_row2[2] = FlexCell("")
header_row2[3] = FlexCell("")


header_row3[1] = FlexCell(pot("In-Area Employment Efficiency (Primary Jobs)", body_header_text))
header_row3[2] = FlexCell("")
header_row3[3] = FlexCell("")


header_row4[1] = FlexCell(pot("Outflow Job Characteristics", body_header_text))
header_row4[2] = FlexCell("")
header_row4[3] = FlexCell("")



io_table1 <- addHeaderRow(io_table1, header_row)
io_table1 <- addFooterRow(io_table1, value = "", colspan=3)
io_table2 <- addHeaderRow(io_table2, header_row2)
io_table2 <- addFooterRow(io_table2, value = "", colspan=3)
io_table3 <- addHeaderRow(io_table3, header_row3)
io_table3 <- addFooterRow(io_table3, value = "", colspan=3)
io_table4 <- addHeaderRow(io_table4, header_row4)
io_table4 <- addFooterRow(io_table4, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                           text.properties=source_text)


io_table1 <- setFlexTableBorders(io_table1
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

io_table2 <- setFlexTableBorders(io_table2
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

io_table3 <- setFlexTableBorders(io_table3
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

io_table4 <- setFlexTableBorders(io_table4
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)


io_table1 <- setFlexTableWidths(io_table1,c(4.5,1,1))
io_table2 <- setFlexTableWidths(io_table2,c(4.5,1,1))
io_table3 <- setFlexTableWidths(io_table3,c(4.5,1,1))
io_table4 <- setFlexTableWidths(io_table4,c(4.5,1,1))



io_table1[,1,side="left"] <- borderProperties(style='solid')
io_table1[,3,side="right"] <- borderProperties(style='solid')
io_table1[1,1,to='header', side="left"] <- borderProperties(style='solid')
io_table1[1,3,to='header', side="right"] <- borderProperties(style='solid')
io_table1[1,, to='header', side="top"] <- borderProperties(style='solid')
io_table1[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
io_table1[to = 'footer', side = 'top'] <- borderProperties(style='none')


io_table2[,1,side="left"] <- borderProperties(style='solid')
io_table2[,3,side="right"] <- borderProperties(style='solid')
io_table2[1,1,to='header', side="left"] <- borderProperties(style='solid')
io_table2[1,3,to='header', side="right"] <- borderProperties(style='solid')
io_table2[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
io_table2[to = 'footer', side = 'top'] <- borderProperties(style='none')


io_table3[,1,side="left"] <- borderProperties(style='solid')
io_table3[,3,side="right"] <- borderProperties(style='solid')
io_table3[1,1,to='header', side="left"] <- borderProperties(style='solid')
io_table3[1,3,to='header', side="right"] <- borderProperties(style='solid')
io_table3[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
io_table3[to = 'footer', side = 'top'] <- borderProperties(style='none')


io_table4[,1,side="left"] <- borderProperties(style='solid')
io_table4[,3,side="right"] <- borderProperties(style='solid')
io_table4[1,1,to='header', side="left"] <- borderProperties(style='solid')
io_table4[1,3,to='header', side="right"] <- borderProperties(style='solid')
io_table4[to = 'footer', side = 'right'] <- borderProperties(style='none')
io_table4[to = 'footer', side = 'left'] <- borderProperties(style='none')
io_table4[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
io_table4[to = 'footer', side = 'top'] <- borderProperties(style='solid')




io_table1[, 1] <- left_align
io_table1[,2:3] <- right_align


io_table2[, 1] <- left_align
io_table2[,2:3] <- right_align

io_table3[, 1] <- left_align
io_table3[,2:3] <- right_align

io_table4[, 1] <- left_align
io_table4[,2:3] <- right_align



ls <- addFlexTable(ls, io_table1)
ls <- addFlexTable(ls, io_table2)
ls <- addFlexTable(ls, io_table3)
ls <- addFlexTable(ls, io_table4)

#-----SECOND PAGE OF INFLOW/OUTFLOW SUMMARY 

ls <- addPageBreak(ls)
ls <- addParagraph(ls, value = area_title, stylename = 'Normal', par.properties = parProperties(text.align='left', padding.bottom = 8))
ls <- addParagraph(ls, value = pot('Labor Inflow/Outflow Summary (continued)', title_format), stylename = 'Normal', 
                   par.properties = parProperties(text.align='left', padding.top = 8))


io_table5 <- FlexTable(in_out_summ[20:29,], body.text.props = body_text, header.columns = FALSE)
io_table6 <- FlexTable(in_out_summ[30:39,], body.text.props = body_text, header.columns = FALSE)



header_row5 = FlexRow()
header_row6 = FlexRow()



header_row5[1] = FlexCell(pot("Inflow Job Characteristics (Primary Jobs)", body_header_text))
header_row5[2] = FlexCell(pot("Count", body_header_text),
                         par.properties = parProperties(text.align='right'))

header_row5[3] = FlexCell(pot("Share", textProperties(font.weight = 'bold')),
                         par.properties = parProperties(text.align='right'))


header_row6[1] = FlexCell(pot("Interior Flow Job Characteristics (Primary Jobs)", body_header_text))
header_row6[2] = FlexCell("")
header_row6[3] = FlexCell("")


io_table5 <- addHeaderRow(io_table5, header_row5)
io_table5 <- addFooterRow(io_table5, value = "", colspan=3)
io_table6 <- addHeaderRow(io_table6, header_row6)
io_table6 <- addFooterRow(io_table6, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                           text.properties=source_text)


io_table5 <- setFlexTableBorders(io_table5
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)

io_table6 <- setFlexTableBorders(io_table6
                                  ,inner.vertical = no_border
                                  ,inner.horizontal = no_border
                                  ,outer.vertical = no_border
                                  ,outer.horizontal = no_border)




io_table5 <- setFlexTableWidths(io_table5,c(4.5,1,1))
io_table6 <- setFlexTableWidths(io_table6,c(4.5,1,1))


io_table5[,1,side="left"] <- borderProperties(style='solid')
io_table5[,3,side="right"] <- borderProperties(style='solid')
io_table5[1,1,to='header', side="left"] <- borderProperties(style='solid')
io_table5[1,3,to='header', side="right"] <- borderProperties(style='solid')
io_table5[1,, to='header', side="top"] <- borderProperties(style='solid')
io_table5[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
io_table5[to = 'footer', side = 'top'] <- borderProperties(style='none')


io_table6[,1,side="left"] <- borderProperties(style='solid')
io_table6[,3,side="right"] <- borderProperties(style='solid')
io_table6[1,1,to='header', side="left"] <- borderProperties(style='solid')
io_table6[1,3,to='header', side="right"] <- borderProperties(style='solid')
io_table6[to = 'footer', side = 'right'] <- borderProperties(style='none')
io_table6[to = 'footer', side = 'left'] <- borderProperties(style='none')
io_table6[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
io_table6[to = 'footer', side = 'top'] <- borderProperties(style='solid')


io_table5[, 1] <- left_align
io_table5[,2:3] <- right_align


io_table6[, 1] <- left_align
io_table6[,2:3] <- right_align


ls <- addFlexTable(ls, io_table5)
ls <- addFlexTable(ls, io_table6)



#---------------------------------------------------------------------------------------------------------
#                                      WORKER DESTINATION SUMMARY
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)
ls <- addTitle(ls, 'Work Destination Summary', level=1)
ls <- addParagraph(ls, pot(paste('Where Workers are Employed Who Live in', area_name, sep = " "), title_format), 
                   par.properties = parProperties(text.align='left', padding.top = 8))

wd_table1 <- FlexTable(work_dest_final[1,], body.text.props = body_text, header.columns = FALSE)
wd_table2 <- FlexTable(work_dest_final[2:27,], body.text.props = body_text, header.columns = FALSE)



header_row = FlexRow()
header_row2 = FlexRow()



header_row[1] = FlexCell("")
header_row[2] = FlexCell(pot("Count", body_header_text),
                          par.properties = parProperties(text.align='right'))

header_row[3] = FlexCell(pot("Share", body_header_text),
                          par.properties = parProperties(text.align='right'))


header_row2[1] = FlexCell(pot("Job Counts by Counties Where Workers are Employed - Primary Jobs", body_header_text))
header_row2[2] = FlexCell("")
header_row2[3] = FlexCell("")


wd_table1 <- addHeaderRow(wd_table1, header_row)
wd_table1 <- addFooterRow(wd_table1, value = "", colspan=3)
wd_table2 <- addHeaderRow(wd_table2, header_row2)
wd_table2 <- addFooterRow(wd_table2, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                          text.properties=source_text)


wd_table1 <- setFlexTableBorders(wd_table1
                                 ,inner.vertical = no_border
                                 ,inner.horizontal = no_border
                                 ,outer.vertical = no_border
                                 ,outer.horizontal = no_border)

wd_table2 <- setFlexTableBorders(wd_table2
                                 ,inner.vertical = no_border
                                 ,inner.horizontal = no_border
                                 ,outer.vertical = no_border
                                 ,outer.horizontal = no_border)




wd_table1 <- setFlexTableWidths(wd_table1,c(4.5,1,1))
wd_table2 <- setFlexTableWidths(wd_table2,c(4.5,1,1))


wd_table1[1,1] = body_header_text

wd_table1[,1,side="left"] <- borderProperties(style='solid')
wd_table1[,3,side="right"] <- borderProperties(style='solid')
wd_table1[1,1,to='header', side="left"] <- borderProperties(style='solid')
wd_table1[1,3,to='header', side="right"] <- borderProperties(style='solid')
wd_table1[1,, to='header', side="top"] <- borderProperties(style='solid')
wd_table1[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wd_table1[to = 'footer', side = 'top'] <- borderProperties(style='none')


wd_table2[,1,side="left"] <- borderProperties(style='solid')
wd_table2[,3,side="right"] <- borderProperties(style='solid')
wd_table2[1,1,to='header', side="left"] <- borderProperties(style='solid')
wd_table2[1,3,to='header', side="right"] <- borderProperties(style='solid')
wd_table2[to = 'footer', side = 'right'] <- borderProperties(style='none')
wd_table2[to = 'footer', side = 'left'] <- borderProperties(style='none')
wd_table2[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
wd_table2[to = 'footer', side = 'top'] <- borderProperties(style='solid')


wd_table1[, 1] <- left_align
wd_table1[,2:3] <- right_align


wd_table2[, 1] <- left_align
wd_table2[,2:3] <- right_align


ls <- addFlexTable(ls, wd_table1)
ls <- addFlexTable(ls, wd_table2)




#---------------------------------------------------------------------------------------------------------
#                                      HOME DESTINATION SUMMARY
#---------------------------------------------------------------------------------------------------------

ls <- addPageBreak(ls)
ls <- addTitle(ls, 'Home Destination Summary', level=1)
ls <- addParagraph(ls, value = pot(paste('Where Workers Live Who are Employed in', area_name, sep=" "), title_format), stylename = 'Normal', 
                   par.properties = parProperties(text.align='left', padding.top = 8))

hd_table1 <- FlexTable(home_dest_final[1,], body.text.props = body_text, header.columns = FALSE)
hd_table2 <- FlexTable(home_dest_final[2:27,], body.text.props = body_text, header.columns = FALSE)



header_row = FlexRow()
header_row2 = FlexRow()



header_row[1] = FlexCell("")
header_row[2] = FlexCell(pot("Count", body_header_text),
                         par.properties = parProperties(text.align='right'))

header_row[3] = FlexCell(pot("Share", body_header_text),
                         par.properties = parProperties(text.align='right'))


header_row2[1] = FlexCell(pot("Job Counts by Counties Where Workers Live - Primary Jobs", body_header_text))
header_row2[2] = FlexCell("")
header_row2[3] = FlexCell("")


hd_table1 <- addHeaderRow(hd_table1, header_row)
hd_table1 <- addFooterRow(hd_table1, value = "", colspan=3)
hd_table2 <- addHeaderRow(hd_table2, header_row2)
hd_table2 <- addFooterRow(hd_table2, value = "Source: U.S. Census Bureau, OnTheMap Application and Longitudinal Employer-Household Dynamics program.", colspan=3,
                          text.properties=source_text)


hd_table1 <- setFlexTableBorders(hd_table1
                                 ,inner.vertical = no_border
                                 ,inner.horizontal = no_border
                                 ,outer.vertical = no_border
                                 ,outer.horizontal = no_border)

hd_table2 <- setFlexTableBorders(hd_table2
                                 ,inner.vertical = no_border
                                 ,inner.horizontal = no_border
                                 ,outer.vertical = no_border
                                 ,outer.horizontal = no_border)




hd_table1 <- setFlexTableWidths(hd_table1,c(4.5,1,1))
hd_table2 <- setFlexTableWidths(hd_table2,c(4.5,1,1))


hd_table1[1,1] = body_header_text

hd_table1[,1,side="left"] <- borderProperties(style='solid')
hd_table1[,3,side="right"] <- borderProperties(style='solid')
hd_table1[1,1,to='header', side="left"] <- borderProperties(style='solid')
hd_table1[1,3,to='header', side="right"] <- borderProperties(style='solid')
hd_table1[1,, to='header', side="top"] <- borderProperties(style='solid')
hd_table1[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
hd_table1[to = 'footer', side = 'top'] <- borderProperties(style='none')


hd_table2[,1,side="left"] <- borderProperties(style='solid')
hd_table2[,3,side="right"] <- borderProperties(style='solid')
hd_table2[1,1,to='header', side="left"] <- borderProperties(style='solid')
hd_table2[1,3,to='header', side="right"] <- borderProperties(style='solid')
hd_table2[to = 'footer', side = 'right'] <- borderProperties(style='none')
hd_table2[to = 'footer', side = 'left'] <- borderProperties(style='none')
hd_table2[to = 'footer', side = 'bottom'] <- borderProperties(style='none')
hd_table2[to = 'footer', side = 'top'] <- borderProperties(style='solid')


hd_table1[, 1] <- left_align
hd_table1[,2:3] <- right_align


hd_table2[, 1] <- left_align
hd_table2[,2:3] <- right_align



ls <- addFlexTable(ls, hd_table1)
ls <- addFlexTable(ls, hd_table2)

#---------------------------------------------------------------------------------------------------------
#                                      ADD END PAGE
#---------------------------------------------------------------------------------------------------------
ls <- addPageBreak(ls)
ls <- addDocument(ls, "labor_shed_end.docx")


#---------------------------------------------------------------------------------------------------------
#                                      CREATE FILE
#---------------------------------------------------------------------------------------------------------

writeDoc(ls, file = paste0('T:/Census Data Center/On-The-Map/Automation/',area_name,' Labor Shed.docx'))
