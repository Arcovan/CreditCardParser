# This program converts credit card statements to CSV-format to import into accounting system
# PDF provided by CC-Company is converted to CSV
# options(encoding = "ISO-8859-1")
# install.packages("dplyr")
# library(dplyr)
# install.packages("devtools")
# devtools::install_github("Arcovan/Rstudio")
# Date : 12 May 2023
# HSBC is work in progress 

library(pdftools)
# ===== Define Functions --------------------------------------------------
ConvertAmount <- function(x) {
  # Remove period separators and convert comma to decimal point
  x <- gsub(".", "", x, fixed = TRUE)  # nb should be "\\." but that does not work I dont knwo why but this works
  x <- gsub(",", ".", x, fixed = TRUE)
  result<- as.numeric(x) # since txt string has format "##.###,##"
  return(result)
} # since txt string has format "##.###,##"
CheckDocType <- function(x) {
  DocType<-"UNKNOWN"
  ProfileTXT<-c("Mastercard|International Card Services BV","Statement ING Corporate Card","www.business.hsbc.fr")
  if (length(grep(ProfileTXT[1], x[1]))!=0) { DocType<-"ICS"}
  if (length(grep(ProfileTXT[2], x[1]))!=0) { DocType<-"ING"}
  if (length(grep(ProfileTXT[3], x[4]))!=0) { DocType<-"HSBC"}
  return(list(DocType = DocType, ProfileTXT = ProfileTXT))
}
Subtype <- function(x) {
  if (length(grep("Mastercard Business Card",x))!=0) {
    Sub<-"BC"} # Business Card
  else { Sub<-"PC" } # Private Card
  return(Sub)
} # for mastercard only can be business or personal

# ==== Read and check Import file ------------------------------------------------------
ifile <- file.choose()
if (!(grepl(".pdf",basename(ifile),ignore.case = TRUE))) {     # Grepl = grep logical
  stop("Please choose file with extension 'pdf'.\n")
}
if (ifile == "") {
  stop("Empty File name [ifile]\n")
}
ofile <- sub(".pdf","-PYukiR.csv", ifile,ignore.case = TRUE) #output file same name but different extension
setwd(dirname(ifile))     #set working directory to input directory where file is
message("Input fileConvertAmount ", basename(ifile), "\nOutput file: ", basename(ofile)) # display filename and output file with full dir name
message("Output file to directory: ", getwd())

# ==== SET Environment ====
options(OutDec = ".")     #set decimal point to "."

# Read PDF--------------------
PDFRaw <- pdftools::pdf_text(ifile)     # read PDF and store in type Vector of Char every page is element
PDFRaw <- gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", PDFRaw, perl = TRUE) # strip double spaces
NrOfPages <- length(PDFRaw)             # number of objects in list = nr of pages in PDF
PDFList <- strsplit(PDFRaw, split = "\n") #type list; page per entry
NROF_RawLines <- sum(lengths(PDFList))
message("Number of lines read: ", NROF_RawLines, " from " , NrOfPages , " page(s).")
if (NROF_RawLines == 0) {
  stop("No characters found in ",basename(ifile)," \nProbably only scanned images in PDF and not a native PDF.")
}
# Add pages to vector of character ------------------------------------------
# for easy processing 
# CCRaw contains full PDF line per line in txt format Every line is an element in the list
CCRaw<-unlist(PDFList)
# # check document type (little too much evaluation but don't know --------
result<-CheckDocType(CCRaw)  #doctype means which credit card supplier / multiple  variables returned
  DocType<-result$DocType
  ProfileTXT<-result$ProfileTXT
if (DocType=="UNKNOWN"){
  message("Either one of the Following keywords were not found:\n")
  m<-c(1:length(ProfileTXT))  #defined in function CheckDoctype
  for (i in m) {
    message(ProfileTXT[i]) 
  }
  stop("Documenttype is not recognised. (No ING and no ICS)")
} 

if (DocType=="ING") { 
  message("ING AFSCHRIFT Recognised in English language.\n")
  DateLineNR<-which(regexpr("\\d{2}\\-\\d{2}-\\d{4}", CCRaw, perl = TRUE)>0)[1]
  LineElements <- unlist(strsplit(CCRaw[DateLineNR], " "))
  DatumAfschrift <-as.Date(LineElements[1],"%d-%m-%Y")
  maand<-strftime(DatumAfschrift,"%m")
  year<-strftime(DatumAfschrift,"%Y")
  Afschriftnr <- LineElements[5]
  PaymentLineNR<-which(regexpr("Payments", CCRaw, perl = TRUE)>0)[1]
  LineElements <- unlist(strsplit(CCRaw[PaymentLineNR], " "))
  Payment<-ConvertAmount(paste(LineElements[4]))
  AccountLineNR<-which(regexpr("Account number", CCRaw, perl = TRUE)>0)[1]+2
  LineElements <- unlist(strsplit(CCRaw[AccountLineNR], " "))
  CCAccount<-LineElements[length(LineElements)-2]
  message("CreditCard Account:", CCAccount, " DatumAfschrift: ", LineElements[1]," Afschrift: ", Afschriftnr)
} #Extract Date Document and other document header info
if (DocType=="ICS") {
  #Extract Date Document and other document header info 
  DateLineNR <- grep("Datum ICS",CCRaw)[1]+1 # find line nr containing Document Date on first page
  LineElements <- unlist(strsplit(CCRaw[DateLineNR], " ")) #split words in line in separate elements
  year <- LineElements[3]
  MonthNL<-c("januari", "februari", "maart", "april","mei", "juni","juli","augustus","september","oktober","november","december")
  DatumAfschrift <- paste(LineElements[1], month.abb[which(LineElements[2]==MonthNL)], year) # dd mmm yyyy
  DatumAfschrift <-as.Date(DatumAfschrift,"%d %b %Y") # DOES NOT WORK WITH DUTCH DATES
  maand<-strftime(DatumAfschrift,"%m")
  
  CCAccount <-LineElements[4]     # Hard coded, could be better.. ICS Klantnummer
  Afschriftnr <- LineElements[5]  # currently not really used ; AFSCHRIFT is composed of YearMonth.
  message("Datum Afschrift: ", paste(LineElements[1], month.abb[which(LineElements[2]==MonthNL)], year), " \nVolgnummer: ", Afschriftnr)
  message("Account:", CCAccount, " [",Subtype(CCRaw), "]")
  # Define 4 digit credit card number and store in Card4DigitsLineNR ---------------------------
  # sometimes multiple cards are reported on 1 PDF and identified by 4 digits
  if (Subtype(CCRaw)=="BC") {    #Business Card
    Card4DigitsLineNR <- grep("Card-nummer:", CCRaw)
  }
  if (Subtype(CCRaw)=="PC") {  #Private Card
    Card4DigitsLineNR <-grep("Uw Card met als laatste vier cijfers", CCRaw)   # NB if more cards on 1 statement statement is NOT split
  }
  if (length(c(Card4DigitsLineNR))>1){message("Multiple Cards:")}
  CCnr<-vector()   #define as vector
  m<-c(1:length(Card4DigitsLineNR))
  for (i in m) {
    LineElements <- unlist(strsplit(CCRaw[Card4DigitsLineNR[i]], " ")) #re-use same variable LineElements for other line
    CCnr<-c(CCnr,LineElements[which(regexpr("\\d{4}",LineElements)>0)[1]]) # credit Card number ****dddd
    message("Kaart:", CCnr[i])
  }
} #Extract Date Document and other document header info 
if (DocType=="HSBC") {
  DateLineNR<-grep("SOLDE DE FIN DE PERIODE", CCRaw)
  DatumAfschrift <-regmatches(CCRaw[DateLineNR], regexpr("\\d{2}\\.\\d{2}\\.\\d{4}",CCRaw[DateLineNR]))   # Extract date from string
  DatumAfschrift <-as.Date(DatumAfschrift,"%d.%m.%Y") 
  maand<-format(DatumAfschrift,"%m")
  year<-format(DatumAfschrift,"%Y")
  #
  IBANLineNR<-grep("IBAN", CCRaw)
  CCRaw[IBANLineNR]
  IBAN <- regmatches(CCRaw[IBANLineNR], regexpr("IBAN [A-Z]{2}\\d{2}\\s(?:\\d{4}\\s?){4}\\d{3}", CCRaw[IBANLineNR]))
  IBAN <- gsub("IBAN ", "", IBAN) # Remove the "IBAN " prefix from the extracted string
  IBAN <- gsub(" ","",IBAN)       # Remove spacesfrom the extracted string
  }
  
pyear<-as.character(as.numeric(year)-1)   # previous year to handle transaction around year end
nyear<-as.character(as.numeric(year)+1)   # next year to handle transaction around year end
#create new empty Matrix mCreditCard
mCreditCard <-
  matrix(data = "",
         nrow = NROF_RawLines,
         ncol = 8) #create empty matrix Import format Yuki Bank Statements
colnames(mCreditCard) <-
  c(
    "IBAN",
    "Valuta",
    "Afschrift",
    "Datum",
    "Rentedatum",
    "Tegenrekening",
    "Naam tegenrekening",
    "Omschrijving"
  )
mCreditCard[, "Omschrijving"] <- CCRaw     # assign vector to column in matrix to start splitting to columns
#create seperate Data Frame since amount is different datatype and does not match matrix
mBedrag <- matrix(data = 0,
                  nrow = NROF_RawLines,
                  ncol = 1)
colnames(mBedrag) <- c("Bedrag")
BedragDF <- data.frame(mBedrag) #convert matrix to data frame
message("Start split ",DocType, "-document in 8 columns and: ", NROF_RawLines, " raw lines.\n")
i <- 1
if (DocType=="ICS") {
  while (i <= NROF_RawLines) {
    CCRegel <- mCreditCard[i, "Omschrijving"] # CCRegel Used for stripping step by step
    Lengte <- nchar(CCRegel)
    PosBIJAF <- regexpr("Bij|Af|credit|debet", CCRegel)[1]  # not perfect with multiple occurences
    LenBIJAF <- attr(regexpr("Bij|Af|credit|debet", CCRegel), "match.length")
    if (PosBIJAF >= 0) {    # BIJ AF found means line relevant for output
      BIJAF <- substr(CCRegel, PosBIJAF, PosBIJAF + LenBIJAF) # text BIJ or Text Af end of line
      #does not work properly with multiple instances of debet etc
      CCRegel <- substr(CCRegel, 1, PosBIJAF - 2) #strip BIJ/AF from text aan einde van regel
      DATSTART <- regexpr("\\d{2}\\s...\\s\\d{2}", CCRegel)[1] #search where (and if) date starts dd mmm dd
      if (DATSTART >= 0) {
        mCreditCard[i, "Datum"] <- paste(substr(CCRegel, DATSTART, 6), year) # dd mmm yyyy Column 4 Datum take part from line and add Year
        CCRegel <- substr(CCRegel, 8, nchar(CCRegel)) #strip datum
        mCreditCard[i, "Rentedatum"] <- paste(substr(CCRegel, 1, 6), year) #Column 5 rente datum
        CCRegel <- substr(CCRegel, 8, nchar(CCRegel)) #strip datum 2
        PosBedrag <-   #start positie amount
          regexpr("(\\d{1,3}(\\.?\\d{3})*(,\\d{2})?$)|(^\\d{1,3}(,?\\d{3})*(\\.\\d{2})?$)", CCRegel)[1]# check bot dec point as dec comma
        sBedrag <- substr(CCRegel, PosBedrag, nchar(CCRegel)) # Column 9
        if (regexpr("Af|debet", BIJAF)[1] > 0) {
          # af or credit
          BedragDF$Bedrag[i] <- -ConvertAmount(paste(sBedrag))
        }
        else {
          BedragDF$Bedrag[i] <- ConvertAmount(paste(sBedrag))
        }
        mCreditCard[i, "Naam tegenrekening"] <- substr(CCRegel, 1, PosBedrag - 6) #Column 7
        mCreditCard[i, "Omschrijving"] <- substr(CCRegel,1, PosBedrag-2) #after all stripping Description is left
      }
    }
    i <- i + 1
  }

  if (length(CCnr)==1){
    message("Card: ",CCnr[1],"/",Card4DigitsLineNR[1],"/","/TOT:",NROF_RawLines)
    mCreditCard[(Card4DigitsLineNR[1]+1):(NROF_RawLines),"Omschrijving"]<-paste(CCnr[1],":",mCreditCard[(Card4DigitsLineNR[1]+1):(NROF_RawLines),"Omschrijving"]) #add 4 digits cc to description
  } else {
    if (length(CCnr)>1){
      message("Cards: ",CCnr[1],"/Van:",Card4DigitsLineNR[1]+1," TOT:",Card4DigitsLineNR[2]-1)
      mCreditCard[(Card4DigitsLineNR[1]+1):(Card4DigitsLineNR[2]-1),"Omschrijving"]<-paste(CCnr[1],":",mCreditCard[(Card4DigitsLineNR[1]+1):(Card4DigitsLineNR[2]-1),"Omschrijving"])
      message("Cards: ",CCnr[2],"/Van:",Card4DigitsLineNR[2]," TOT:",NROF_RawLines)
      mCreditCard[(Card4DigitsLineNR[2]+1):(NROF_RawLines),"Omschrijving"]<-paste(CCnr[2],":",mCreditCard[(Card4DigitsLineNR[2]+1):(NROF_RawLines),"Omschrijving"])
    }
  } # Add credit card 4 digit number to lines
}
if (DocType=="ING") {
  while (i <= NROF_RawLines) {
    CCRegel <- mCreditCard[i, "Omschrijving"] # CCRegel Used for stripping step by step
    Lengte <- nchar(CCRegel)
    DATSTART <- regexpr("\\d{2}\\s...\\s.", CCRegel)[1] #search where (and if) date starts dd mmm dd
    if (DATSTART >= 0) {
      mCreditCard[i, "Datum"] <- paste(substr(CCRegel, DATSTART, 6), year) # dd mmm yyyy Column 4 Datum take part from line and add Year
      mCreditCard[i, "Rentedatum"] <- mCreditCard[i, "Datum"]#Column 5 rente datum
      CCRegel <- substr(CCRegel, 8, nchar(CCRegel)) #strip datum
      PosBedrag <-   #start positie amount at end of string
        regexpr("(\\d{1,3}(\\.?\\d{3})*(,\\d{2})?$)|(^\\d{1,3}(,?\\d{3})*(\\.\\d{2})?$)", CCRegel)[1] # check both dec point as dec comma at end of string
      sBedrag <- substr(CCRegel, PosBedrag, nchar(CCRegel)) # Column 9
      BedragDF$Bedrag[i] <- -ConvertAmount(paste(sBedrag))
      if (Payment==BedragDF$Bedrag[i]) {
        BedragDF$Bedrag[i] <- -BedragDF$Bedrag[i]
      }
      NaamTR<-paste(unlist(strsplit(CCRegel, " "))[1])
      mCreditCard[i, "Naam tegenrekening"] <- NaamTR #Column 7
    }
    mCreditCard[i, "Omschrijving"] <- CCRegel #after all stripping Description is left
    i <- i + 1
  }
}
CreditCardDF <- as.data.frame(mCreditCard)
CreditCardDF <- cbind.data.frame(CreditCardDF, BedragDF)
View(mCreditCard)
# Delete empty Lines
if (length(which(CreditCardDF$Bedrag == 0)) != 0) {
  CreditCardDF <-
    CreditCardDF[-c(which(CreditCardDF$Bedrag == 0)),] 
}  #delete row with Bedrag == 0
if (length(which(CreditCardDF$Omschrijving == "")) != 0) {
  CreditCardDF <-
    CreditCardDF[-c(which(CreditCardDF$Omschrijving == "")),]
}  #delete row with empty omschrijving
#
CreditCardDF$Valuta <- "EUR"
# create Afschiftnummer
CreditCardDF$Afschrift <- paste(year, maand, sep = "") # yyyymm bepaal volgnummer maand
if (DocType=="ICS") {
  CreditCardDF$`Naam tegenrekening`<-gsub("GEINCASSEERD VORIG S",'MasterCard',CreditCardDF$`Naam tegenrekening`)
  CreditCardDF$IBAN <- CCAccount # should change per card now all have same creditcard number
  if (Subtype(CCRaw)=="PC") {CreditCardDF$IBAN <- CCAccount}
}
if (DocType=="ING") {
  CreditCardDF$IBAN <- CCAccount
  Line<-grep("Amount will be debited around",CreditCardDF$Omschrijving)
  if (Line>0) {CreditCardDF <- CreditCardDF[-(Line),]}    # Delete line with total amount
  Line<-grep("PAYMENT RECEIVED",CreditCardDF$Omschrijving)
  if (Line>0) {CreditCardDF$Bedrag[Line]<-(-CreditCardDF$Bedrag[Line])} # Positive amount is direct debet of previous account
}
# prepare for date conversion
CreditCardDF$Datum <- gsub("mrt", 'mar', CreditCardDF$Datum) # mrt is not recognised as month for date conversion so replace
CreditCardDF$Datum <- gsub("mei", 'may', CreditCardDF$Datum) # mei is not recognised as month for date conversion so replace
CreditCardDF$Datum <- gsub("okt", 'oct', CreditCardDF$Datum) # okt is not recognised as month for date conversion so replace
CreditCardDF$Datum <- as.Date(CreditCardDF$Datum, "%d %b %Y")
CreditCardDF$Datum <- format(CreditCardDF$Datum, "%d-%m-%Y")
CreditCardDF$Rentedatum <- gsub("mrt", 'mar', CreditCardDF$Rentedatum)
CreditCardDF$Rentedatum <- gsub("mei", 'may', CreditCardDF$Rentedatum)
CreditCardDF$Rentedatum <- gsub("okt", 'oct', CreditCardDF$Rentedatum)
CreditCardDF$Rentedatum <- as.Date(CreditCardDF$Rentedatum, "%d %b %Y")
CreditCardDF$Rentedatum <- format(CreditCardDF$Rentedatum, "%d-%m-%Y")
# some magic to change the year of transaction that were in previous year compared to Date of statement since line do not have year indication
if (as.numeric(maand)<=2) {
  CreditCardDF$Datum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Datum,"%m")))>1)]<-sub(year, pyear,CreditCardDF$Datum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Datum,"%m")))>1)])
  CreditCardDF$Rentedatum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Rentedatum,"%m")))>1)]<-sub(year, pyear,CreditCardDF$Rentedatum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Rentedatum,"%m")))>1)])
} # statement jan or feb with possible lines from previous year
if (as.numeric(maand)>=11) {
  CreditCardDF$Datum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Datum,"%m")))>1)]<-sub(year, nyear,CreditCardDF$Datum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Datum,"%m")))>1)])
  CreditCardDF$Rentedatum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Rentedatum,"%m")))>1)]<-sub(year, pyear,CreditCardDF$Rentedatum[which(abs(as.numeric(maand)-as.numeric(strftime(CreditCardDF$Rentedatum,"%m")))>1)])
} # unlikely to have transactions later then statement date but just in case
# end magic
View(CreditCardDF)
summary(CreditCardDF$Bedrag)
#write.csv2(CreditCardDF, file = ofile, row.names = FALSE)         # Delimeter ; no row numbers
write.table(
  CreditCardDF,
  file = ofile,   #Output file define at start
  quote = FALSE,
  sep = ";",
  dec = ".",
  row.names = FALSE,
  #col.names = ColumnNames  # Vanwege de underscore in de header die geen spatie kan zijn
)
address<-as.data.frame(CreditCardDF$`Naam tegenrekening`)
write.table(address, file = "address.txt", append = TRUE,
            row.names = FALSE, col.names = FALSE)
source("/Users/arco/Dropbox/R-Studio/MasterCardPDF/InformUser.R") # print summary
# compositie regular expression
#(?:^           # beginning of string
#     \d{1,3}      # one, two, or three digits
#   (?:
#       \.?        # optional separating period
#       \d{3}      # followed by exactly three digits
#   )*           # repeat this subpattern (.###) any number of times (including none at all)
#     (?:,\d{2})?  # optionally followed by a decimal comma and exactly two digits
#     $)             # End of string.
# |              # ...or...
#   (?:^           # beginning of string
#      \d{1,3}      # one, two, or three digits
#    (?:
#        ,?         # optional separating comma
#        \d{3}      # followed by exactly three digits
#    )*           # repeat this subpattern (,###) any number of times (including none at all)
#      (?:\.\d{2})? # optionally followed by a decimal perioda and exactly two digits
#      $)             # End of string.
