# This program converts creditcard statements to format to import in accounting system
# PDF provided by CC-Company is converted to CSV
# options(encoding = "ISO-8859-1")
# getOption("encoding")
# rm(list=ls())
# getOption("OutDec")        # check what decimal point is and return "." or ","
#install.packages("devtools")
#devtools::install_github("Arcovan/Rstudio")
ConvertAmount <- function(x) {
  TxtBedrag <- x
  TxtBedrag <- (sub(",", "d", TxtBedrag, fixed = TRUE))
  TxtBedrag <- (sub(".", "t", TxtBedrag, fixed = TRUE))
  TxtBedrag <- (sub("d", ".", TxtBedrag, fixed = TRUE))
  TxtBedrag <- (sub("t", "", TxtBedrag, fixed = TRUE))
  TxtBedrag <- as.numeric(TxtBedrag)
  return(TxtBedrag)
}
CheckDocType <- function(x) {
  DocType<-"UNKNOWN"
  ProfileTXT<-c("Mastercard|International Card Services BV","Statement ING Corporate Card")
  if (length(grep(ProfileTXT[1], x))!=0) { DocType<-"ICS"}
  if (length(grep(ProfileTXT[2], x))!=0) { DocType<-"ING"}
  return(DocType)
}
Subtype <- function(x) {
  if (length(grep("Mastercard Business Card",x))!=0) {
    Sub<-"BC"} # Business Card
  else { Sub<-"PC" } # Private Card
  return(Sub)
}
options(OutDec = ",")        # set decimal point to ","
ifile <- file.choose()       # Select Importfile
if (grep(".pdf", ifile) < 0) {
  stop("Please choose file with extension 'pdf'.\n")
}
ofile <- sub("pdf","csv", ifile) #output file same name but different extension
message("Input file: ", ifile, "\nOutput file: ", ofile)
setwd(dirname(ifile))         # set working directory to input directory where file is
message("Output file to directory: ", getwd(),"\n")
# Start reading PDF -> PDFList
PDFFILE <- pdftools::pdf_text(ifile) # read PDF and store in type Char
PDFFILE <- gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", PDFFILE, perl = TRUE) # strip double spaces
PDFList <- strsplit(PDFFILE, split = "\n") #type list; page per entry
NrOfPages <- length(PDFList) # number of objects in list = nr of pages in PDF
NROF_RawLines <- sum(lengths(PDFList))
message("Number of lines read: ", NROF_RawLines, " from " , NrOfPages , " page(s).")
if (NROF_RawLines == 0) {
  fname<-sub(dirname(ifile),"File:.", ifile)
  stop("No characters found in ",fname," \nProbably only scanned images in PDF and not a native PDF.")
}
# add all elements of list (pages) to 1 vector for easy processing
page <- 1
CCRaw <- PDFList[[page]]
page <- page + 1
while (page <= NrOfPages) {
  CCRaw <- append(CCRaw, PDFList[[page]]) #all seperate pages to 1 list plain raw PDF 
  page <- page + 1
}
View(CCRaw)
# check document type (little too much evaluation but dont know case statement)
DocType<-CheckDocType(CCRaw)  #doctype means which creditcard supplier
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
  #MasterCard ICS
  DateLineNR <- which(substr(CCRaw, 1, 9) == "Datum ICS")[1] + 1 # find line nr containing Document Date
  LineElements <- unlist(strsplit(CCRaw[DateLineNR], " ")) #split words in line in seperate elements
  year <- LineElements[3]
  MonthNL<-c("januari", "februari", "maart", "april","mei", "juni","juli","augustus","september","oktober","november","december")
  month.abb[which(LineElements[2]==MonthNL)] #translate to fit as.date function
  DatumAfschrift <- paste(LineElements[1], month.abb[which(LineElements[2]==MonthNL)], year) # dd mmm yyyy
  DatumAfschrift <-as.Date(DatumAfschrift,"%d %b %Y") # DOES NOT WORK WITH DUTCH  DATES
  maand<-strftime(DatumAfschrift,"%m")
  CCAccount <-LineElements[4] 
  Afschriftnr <- LineElements[5] #currently not really used ; AFSCHRIFT is composed of YearMonth
  message("DatumAfschrift: ", paste(LineElements[1], month.abb[which(LineElements[2]==MonthNL)], year)," Afschrift: ", Afschriftnr)
  message("Account:", CCAccount, " [",Subtype(CCRaw), "]")
  if (Subtype(CCRaw)=="BC") {    #Business Card
    CCLine <- grep("Card-nummer:", CCRaw)
  }
  if (Subtype(CCRaw)=="PC") {  #Private Card
    CCLine <-grep("Uw Card met als laatste vier cijfers", CCRaw)   # NB if more cards on 1 statement statement is NOT split
  }
  if (length(c(CCLine))>1){message("Multiple Cards:")}
  CCnr<-vector()
  m<-c(1:length(CCLine))
  for (i in m) {
    LineElements <- unlist(strsplit(CCRaw[CCLine[i]], " ")) #re-use same variable LineElements for other line
    CCnr<-c(CCnr,LineElements[which(regexpr("\\d{4}",LineElements)>0)[1]]) # credit Card number ****dddd
    message("Kaart:", CCnr[i])
  }
} #Extract Date Document and other document header info 
#create new empty Matrix
mCreditCard <-
  matrix(data = "",
         nrow = NROF_RawLines,
         ncol = 8) #create empty matrix Import format Yuki Bank Statements
colnames(mCreditCard) <-
  c(
    "IBAN",
    "Valuta",
    "AFSCHRIFT",
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
message("Start split ",DocType, "-document in 8 columns and: ", NROF_RawLines, " lines.\n")
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
      }
      mCreditCard[i, "Omschrijving"] <- CCRegel #after all stripping Description is left
    }
    i <- i + 1
  }
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
View(mCreditCard)
CreditCardDF <- as.data.frame(mCreditCard)
CreditCardDF <- cbind.data.frame(CreditCardDF, BedragDF)
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
CreditCardDF$AFSCHRIFT <- paste(year, maand, sep = "") # yyyymm bepaal volgnummer maand
if (DocType=="ICS") {
  CreditCardDF$IBAN <- CCnr[1] # should change per card now all have same creditcard number
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
View(CreditCardDF)
write.csv2(CreditCardDF, file = ofile, row.names = FALSE)         # Delimeter ; no row numbers
source("/Users/arco/Dropbox/R-Studio/MasterCardPDF/InformUser.R")
message("Export klaar in file:", ofile)
message("Totale uitgaven dit afschrift:", sum(CreditCardDF[CreditCardDF$Bedrag < 0, ]$Bedrag)) #sum only negatove amounts
message("Vorig geincasseerd saldo:", sum(CreditCardDF[CreditCardDF$Bedrag > 0, ]$Bedrag))
