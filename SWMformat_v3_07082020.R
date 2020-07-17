time_to_run <- Sys.time()
#----------------------- activate library -----------------------
library(readxl)
library(stringr)
library(openxlsx)
library(data.table)

#copy files in SWM data cleaning folder and change code filenames to latest files
#----------------------- read files ----------------------------
PRODDUMP <- read_excel("data/SWM Export_SEI ALL 6-29-20.xlsx", 
                                  col_types = c("text","skip","text","text",
                                                "text","text","text","skip",
                                                "skip","text","skip","text",
                                                "skip","date","date","skip",
                                                "text","numeric","numeric","text",
                                                "numeric","text"))
Location_List <- read_excel("data/Current Master Location List.xlsx", sheet = "Store Data",
                                           col_types = c("text","text","text","text",
                                                         "text","skip","skip","skip",
                                                         "skip","text","text","text",
                                                         "text","skip","skip","skip",
                                                         "skip","skip","skip","text",
                                                         "skip","skip","skip","skip",
                                                         "skip","skip","skip","skip",
                                                         "skip","text","skip","skip",
                                                         "skip","skip","skip","skip",
                                                         "skip","skip","text","skip",
                                                         "skip","skip","skip","skip",
                                                         "skip","skip","skip"))
ZIP_CBSA <- read_excel("Data/ZIP-CBSA.xlsx", sheet="ZIPCBSA",
                       col_types = c("skip","text","text","skip","skip","skip"))

#------------------ copy from original -----------------
AllLocations<-Location_List
AllDataSWM<-PRODDUMP
AllCBSA <- ZIP_CBSA

ScheduledtoCoverage<-function (DataSWM,Locations,CBSA) {
#create key & clean data
DataSWM$key <-paste(substr(DataSWM$name,1,5),DataSWM$Store, sep="")
#truncate to 5 digit zips for USA
Locations$Zip<-ifelse(Locations$Country=="CAN",Locations$Zip,substr(Locations$Zip,1,5))

#add closed future date filter
DataSWM$Closed[DataSWM$plan_start_date >as.POSIXct(Sys.Date())]<-"FutureStartDate"
DataSWM$Closed[DataSWM$plan_end_date < as.POSIXct(Sys.Date())]<-"Closed"
#scheduled PM filters - add ACCA to HVAC
DataSWM$key[substr(DataSWM$name,1,5)=="CA AC"]<- paste("HVAC ",DataSWM$Store[substr(DataSWM$name,1,5)=="CA AC"], sep="" )

#creating the coverage dataset - remove closed, inactive & fixed PMS
coverageSP<-DataSWM
coverageSP$Closed<- paste("a",coverageSP$Closed,sep="")
coverageSP$Closed <-as.factor(coverageSP$Closed)
coverageSP$status<-as.factor(coverageSP$status)
coverageSP$line_of_service <-as.factor(coverageSP$line_of_service)
coverageSP$name <-as.factor(coverageSP$name)
coverageSP$site_plan_type <-as.factor(coverageSP$site_plan_type)

#keep only active rows
statusfilter <-coverageSP$status=="Active"
#keep only open rows
closedfilter <- coverageSP$Closed=="aNA"
#keep schedule PM rows + ACCA + Pest + Safes
PMnamefilter <- str_detect(coverageSP$name,"Scheduled") | str_detect(coverageSP$name,"ACCA") | str_detect(coverageSP$name,"Safe") | str_detect(coverageSP$name,"Pest")
#Billable Rows for Safes and Pest
Billablefilter <- str_detect(coverageSP$name,"Fixed Billable")
#keep rows who's start date is not in future
startdatefilter <- coverageSP$plan_start_date>as.POSIXct(Sys.Date())
#apply filters
coverageSP<-coverageSP[statusfilter & closedfilter & PMnamefilter& !startdatefilter & !Billablefilter,]

#create SP Name : Num column
coverageSP$SchedSP<-paste(coverageSP$`SP Number`,coverageSP$`SP Name`,sep=":")
#coverage SP table
coverageSP<-coverageSP[,c("key","SchedSP")]

#merge CBSA with locations
Locations<-merge(Locations,CBSA,by.x="Zip",by.y="ZIPCODE",all.x=TRUE)
#merge SP info
Fulllist <- merge(DataSWM,coverageSP, by="key", all.x=TRUE)
#merge geo info
Fulllist <- merge(Locations,Fulllist, by.y="Store", by.x="Location No", all.y=TRUE)
Fulllist$key=NULL
colnames(Fulllist) <-c("Site Num","ZIP","Zone Num","Zone Name","Market Num","Market Name","City","State","Country","County","StoreType","ServiceCenter","CBSA",
                       "Cust Num","LOS","PM Short Desc","SP Name","SP Num","Plan Type","Status","Plan Start Date","Plan End Date",
                       "Service_Frequency","Provider_Fee","Provider_Annual_Limit","Provider_Billing_Frequency",
                       "Customer_Fee","Customer_Billing_Frequency","Closed?","Coverage SP")
Fulllist<-Fulllist[c("Cust Num","Site Num","StoreType","ServiceCenter","Zone Num","Zone Name","Market Num","Market Name",
                     "City","County","CBSA","State","Country","ZIP",
                     "LOS","PM Short Desc","SP Name","SP Num","Coverage SP","Plan Type","Status","Closed?","Plan Start Date","Plan End Date",
                     "Service_Frequency","Provider_Fee","Provider_Annual_Limit","Provider_Billing_Frequency",
                     "Customer_Fee","Customer_Billing_Frequency")]

return(Fulllist)
#clean other datasets
rm(coverageSP,DataSWM,Locations,statusfilter,closedfilter,PMnamefilter,CBSA)

}

BevBarSWM <- ScheduledtoCoverage(AllDataSWM[str_detect(AllDataSWM$name,"Bev Bar"),],AllLocations,AllCBSA)
NonBevBarSWM <- ScheduledtoCoverage(AllDataSWM[!str_detect(AllDataSWM$name,"Bev Bar"),],AllLocations,AllCBSA)
SWMCoverage<-rbind(BevBarSWM,NonBevBarSWM)

filenameis<-paste("SWMCoverage_Geo+Coverage_",year(Sys.Date()),"_",month(Sys.Date()),".xlsx",sep = "")

write.xlsx(SWMCoverage, filenameis)


print(Sys.time()-time_to_run)
