library('RCurl')
library('rjson')
library(readxl)
df <- read_excel("C:\\Users\\Abhishek\\Downloads\\JPMC\\distance.xlsx")

for(i in 1:10000 ){
  #if(as.character(df[i,1]) != 0)
  #{
  #District <- as.character(df[i,1])
  #}
  print(District)
  place <- as.character(df[i,2])
  A <- "http://maps.googleapis.com/maps/api/distancematrix/json?units=metrics&origins="
  P <- "India"
  C <- "&destinations="
  X <- "&mode=walking"
  #E <-"&Key=AIzaSyC9UbnT_NjU1hFUbJmo5w8XOhqiLoaUP04" #Abhishek
  #E <- "&Key=AIzaSyDBrLWAdjFruRKpCkCljkcqz4q2HcvSyKc" #Chayanika
  E <- "&Key=AIzaSyCMyH-PClfrBf7xPgDSotzX0sBtMI3fsLU" #Ashutosh
  #E <-"&key=AIzaSyDWWgsLTrkybDqGO7QHirQDHWzYCXiqWFY"
  #E<-"&Key=AIzaSyCi3aYuiGgNGFPoW3Q28gSh6K7SaHf-Zns"
  
  #link <- URLencode(paste(A,District," ",P,C,place," ",P,X,E,sep = "")) #Walking
  link <- URLencode(paste(A,District," ",P,C,place," ",P,E,sep = "")) #Driving
  print(place)
  print(i)
  distanceJson <- getURL(link)
  distanceR <- fromJSON(distanceJson)
  if(as.character(distanceR$rows[[1]]$elements[[1]]$status) != "ZERO_RESULTS"  ){
    df[i,3] <- as.character(distanceR$rows[[1]]$elements[[1]]$distance$text)
  }
}

#write.csv(df, "dada.csv")

