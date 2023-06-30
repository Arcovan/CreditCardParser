# This program converts PDF credit card and Bank statements to CSV-format to import into accounting system
# PDF provided by CC-Company is converted to CSV
# Current support PFD statement : CC: Master Card and ING Bank: HSBC
# devtools::install_github("Arcovan/Rstudio")
# Date : 20 May 2023
# HSBC is added but more or less separably; many improvements compare to my first step in R

library(pdftools)
options(OutDec = ".")     #set decimal point to "."
# ===== Define Functions --------------------------------------------------
ConvertAmount <- function(x) {
  # Remove period separators and convert comma to decimal point
  x <- gsub(".", "", x, fixed = TRUE)  # fixed =TRUE treats "." literally and no \\. is required.
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
if (!(grepl("\\.pdf",basename(ifile),ignore.case = TRUE, perl = FALSE))) {     # Grepl = grep logical
  stop("Please choose file with extension 'pdf'.\n")
}
if (ifile == "") {
  stop("Empty File name [ifile]\n")
}
ofile <- sub("\\.pdf","-PYukiR.csv", ifile,ignore.case = TRUE, perl = FALSE) #output file same name but different extension 
setwd(dirname(ifile))     #set working directory to input directory where file is
message("Input  file: ", basename(ifile), "\n") # display filename and output file with full dir name
message("Output file to directory: ", getwd())

# Read PDF--------------------
PDFRaw <- pdftools::pdf_text(ifile)     # read PDF and store in type Vector of Char every page is element
if (!any(nchar(PDFRaw))) {
  stop("No characters found in ",basename(ifile)," \nProbably only scanned images in PDF and not a native PDF.")
}
#PDFRaw <- gsub("\\s{2,}", " ", PDFRaw) # strip double spaces Unclear buf with this way of stripping(does not strip this \n)
PDFRaw <- gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", PDFRaw, perl = TRUE) # strip double spaces (and \n at the end of every page)

# number of objects in list PDFRaw = nr of pages in PDF
PDFList <- strsplit(PDFRaw, split = "\n") #type list; page per entry
CCRaw<-unlist(PDFList) # CCRaw contains full PDF line per line in txt format
NROF_RawLines <- length(CCRaw)
message("Number of lines read: ", NROF_RawLines, " from " , length(PDFRaw) , " page(s).")

# # check document type (little too much evaluation but don't know --------
result<-CheckDocType(CCRaw)  #doctype means which credit card supplier / multiple variables returned
DocType<-result$DocType
ProfileTXT<-result$ProfileTXT
if (DocType=="UNKNOWN"){
  message("Either one of the Following keywords were not found:\n")
  m<-c(1:length(ProfileTXT))  #defined in function CheckDoctype
  for (i in m) {
    message(ProfileTXT[i]) 
  }
  stop("Documenttype is not recognised. (No ING and no ICS and no HSBC bank)")
} # Stop
# read generic header info from CCRaw
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
  ofile <- sub("Rekeningoverzicht", "ICS", ofile, ignore.case = TRUE) # Adjust name to identify CreditCard in Name 
  DateLineNR <- grep("Datum ICS-klantnummer Volgnummer Bladnummer",CCRaw)[1]+1 # find line nr containing Document Date on first page
  LineElements <- unlist(strsplit(CCRaw[DateLineNR], " ")) #split words in line in separate elements
  year <- LineElements[3]
  MonthNL<-c("januari", "februari", "maart", "april","mei", "juni","juli","augustus","september","oktober","november","december")
  DatumAfschrift <- paste(LineElements[1], month.abb[which(LineElements[2]==MonthNL)], year) # dd mmm yyyy
  DatumAfschrift <-as.Date(DatumAfschrift,"%d %b %Y") # DOES NOT WORK WITH DUTCH DATES
  maand<-format(DatumAfschrift,"%m")
  
  Afschriftnr <- LineElements[5]  # currently not really used ; AFSCHRIFT is composed of YearMonth.
  CCAccount <-LineElements[4]     # Hard coded, could be better.. ICS Klantnummer
  
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
  #Extract Date Document and other document header info 
  ofile <- sub("RELEVEDECOMPTE", "HSBC", ofile, ignore.case = TRUE) # Adjust name to identify CreditCard in Name 
  DateLineNR<-grep("SOLDE DE FIN DE PERIODE", CCRaw)
  DatumAfschrift <-regmatches(CCRaw[DateLineNR], regexpr("\\d{2}\\.\\d{2}\\.\\d{4}",CCRaw[DateLineNR]))   # Extract date from string
  DatumAfschrift <-as.Date(DatumAfschrift,"%d.%m.%Y") 
  maand<-format(DatumAfschrift,"%m")
  year<-format(DatumAfschrift,"%Y") #probably not used
  #
  Afschriftnr <- paste(year, maand, sep = "") # 202304 example
  #
  IBANLineNR<-grep("IBAN", CCRaw)
  CCRaw[IBANLineNR]
  IBAN <- regmatches(CCRaw[IBANLineNR], regexpr("(?<=IBAN )[A-Z]{2}[0-9]{2}(?: [0-9]{4}){5} [0-9]{3}", CCRaw[IBANLineNR], perl=TRUE))
  message("Account:", IBAN, " DatumAfschrift: ", DatumAfschrift," Afschrift: ", Afschriftnr)
}#Extract Date Document and other document header info 
message("Output file: ", basename(ofile), "\n")
# New code :create new empty datafram to hold raw statement similar to mCreditCard
StatementDF <- data.frame(Omschrijving = character(length = NROF_RawLines),
                          Bedrag = numeric(length = NROF_RawLines)) # currently Only for HSBC
StatementDF$Omschrijving<-CCRaw # currently Only for HSBC alternative for mCreditCard with only 2 columns
View(StatementDF)
#create new empty Matrix mCreditCard ----
mCreditCard <-
  matrix(data = "",
         nrow = NROF_RawLines,
         ncol = 8) #create empty matrix Import format Yuki Bank Statements (I didnt know DF well enough ;))
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
if (DocType=="ICS") {
  # ------ New style ------------------
  # Extra all amount for line with 'Af' in the staement Line
  NR_StatementLines<-length(grep("^\\b\\d{2} [a-z]{3} \\d{2} [a-z]{3}\\b", CCRaw)) # "dd aaa dd aaa" from the beginning f the line
  YukiDF <- data.frame(IBAN = character(length = NR_StatementLines),
                       Valuta = character(length = NR_StatementLines),
                       Afschrift = integer(length = NR_StatementLines),
                       Datum = character(length = NR_StatementLines),
                       Rentedatum = character(length = NR_StatementLines),
                       Tegenrekening = character(length = NR_StatementLines),
                       Naam_tegenrekening = character(length = NR_StatementLines),
                       Omschrijving = character(length = NR_StatementLines),
                       Bedrag = numeric(NR_StatementLines))
  YukiDF$Omschrijving<- StatementDF$Omschrijving[grepl("^\\b\\d{2} [a-z]{3} \\d{2} [a-z]{3}\\b", StatementDF$Omschrijving)] # "dd aaa dd aaa" from the beginning f the line
  View(YukiDF)
  LineAmount<-regmatches(StatementDF$Omschrijving[grep("Af", StatementDF$Omschrijving)], regexpr("\\b[0-9]{1,3}(?:\\.[0-9]{3})*(?:,[0-9]{2})\\b", StatementDF$Omschrijving[grep("Af", StatementDF$Omschrijving)], perl = TRUE))
  LineAmount<-ConvertAmount(LineAmount)
  StatementDF$Bedrag[grep("Af", StatementDF$Omschrijving)]<--LineAmount
  LineAmount<-regmatches(StatementDF$Omschrijving[grep("Bij", StatementDF$Omschrijving)], regexpr("\\b[0-9]{1,3}(?:\\.[0-9]{3})*(?:,[0-9]{2})\\b", StatementDF$Omschrijving[grep("Bij", StatementDF$Omschrijving)], perl = TRUE))
  LineAmount<-ConvertAmount(LineAmount)
  StatementDF$Bedrag[grep("Bij", StatementDF$Omschrijving)]<-LineAmount
  YukiDF$Omschrijving<-gsub("Af|Bij", "", YukiDF$Omschrijving)
  YukiDF$Omschrijving<-gsub("^\\b\\d{2} [a-z]{3} \\d{2} [a-z]{3}\\b", "", YukiDF$Omschrijving)
  
  # trimws(StatementDF$Omschrijving)
  # regmatches(StatementDF$Omschrijving, regexpr("\\d{2}\\s[a-z]{3}\\s\\d{2}",StatementDF$Omschrijving))
} # New Style using DF (unfinished)

mCreditCard[, "Omschrijving"] <- CCRaw     # assign vector to column in matrix to start splitting to columns
#create seperate Data Frame since amount is different datatype and does not match matrix
BedragDF <- data.frame(Bedrag = numeric(length = NROF_RawLines)) # seperate since a matrix cannot contain both char and numeric 
message("Start split ",DocType, "-raw text in 8 columns and: ", NROF_RawLines, " raw lines.\n")
# extract dates and amount from statement lines
i <- 1
if (DocType=="ICS") {
  while (i <= NROF_RawLines) {
    CCRegel <- mCreditCard[i, "Omschrijving"] # CCRegel Used for stripping step by step
    PosBIJAF <- regexpr("Bij|Af|credit|debet", CCRegel)[1]  # not perfect with multiple occurences
    LenBIJAF <- attr(regexpr("Bij|Af|credit|debet", CCRegel), "match.length")
    # "Bij" or "Af" identifies the real statement lines, from these lines, date and amount are extracted
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
if (DocType=="HSBC") {
  pattern <- "^\\d{2}\\.\\d{2}" # "dd.mm" from the beginning of the line
  # Find rows matching the pattern
  statementlines <- grep(pattern, StatementDF$Omschrijving) # based on date we know it is a statement line
  statementlines
  StatementDF$Omschrijving[statementlines]
  NR_StatementLines <- length(statementlines)
  # Extract the amount using regular expressions
  LineAmount <- regmatches(StatementDF$Omschrijving, regexpr("\\b[0-9]{1,3}(?:\\.[0-9]{3})*(?:,[0-9]{2})\\b", StatementDF$Omschrijving, perl = TRUE))
  LineAmount <- ConvertAmount(LineAmount)  # 2 amount too many and no distinction between debet and credit.
  length(LineAmount)
  LineAmount
  # Initialize an empty vector to store the concatenated descriptions
  concatenated_descriptions <- character(length(statementlines))
  for (i in 1:(length(statementlines) - 1)) {
    start_row <- statementlines[i]
    end_row <- statementlines[i + 1] - 1
    concatenated_description <- paste(StatementDF$Omschrijving[start_row:end_row], collapse = " ")
    concatenated_descriptions[i] <- concatenated_description
  }
  i<-i+1
  start_row <- statementlines[i]
  grep("TOTAL DES MOUVEMENTS",StatementDF$Omschrijving)
  end_row <- grep("TOTAL DES MOUVEMENTS",StatementDF$Omschrijving)-1
  concatenated_description <- paste(StatementDF$Omschrijving[start_row:end_row], collapse = " ")
  concatenated_descriptions[i] <- concatenated_description
  
  Footer<-"---------------------------------------------------------------------------------------------------------------------------------------"
  concatenated_descriptions<-gsub(Footer,"",concatenated_descriptions)
  RIB<-"Relevé d'Identité Bancaire (RIB) RIB Banque Agence N° compte Clé Titulaire : SCI BEACH INVESTMENTS 30056 00228 02282000057 23 Domiciliation : ST TROPEZ Tel : 04 94 55 82 90 IBAN FR76 3005 6002 2802 2820 0005 723 BIC CCFRFRPP"
  concatenated_descriptions<-gsub(RIB,"RIB",concatenated_descriptions, fixed = TRUE) # fixed to avoid as regex
  Footer2<-"HSBC Continental Europe - Société Anonyme au capital de 1 062 332 775 euros - SIREN 775 670 284 RCS Paris Siège social : 38 avenue Kléber 75116 Paris - www.hsbc.fr - N° TVA intracommunautaire FR 70 775 670 284"
  concatenated_descriptions<-gsub(Footer2,"",concatenated_descriptions, fixed = TRUE)
  Header<-"Votre Relevé de Compte e Date Détail des opérations Valeur x Débit Crédit o"
  concatenated_descriptions<-gsub(Header,"",concatenated_descriptions, fixed = TRUE)
  
  YukiDF <- data.frame(IBAN = character(length = NR_StatementLines),
                       Valuta = character(length = NR_StatementLines),
                       Afschrift = integer(length = NR_StatementLines),
                       Datum = character(length = NR_StatementLines),
                       Rentedatum = character(length = NR_StatementLines),
                       Tegenrekening = character(length = NR_StatementLines),
                       Naam_tegenrekening = character(length = NR_StatementLines),
                       Omschrijving = character(length = NR_StatementLines),
                       Bedrag = numeric(NR_StatementLines))
  
  YukiDF$Omschrijving<-concatenated_descriptions
  LineAmount[1:NR_StatementLines+1]
  YukiDF$Bedrag<--LineAmount[1:NR_StatementLines+1] # the first row and the last 2 rows are balance
  YukiDF$IBAN<-gsub("\\s","",IBAN)
  YukiDF$Valuta<-"EUR"
  YukiDF$Afschrift<-Afschriftnr
  YukiDF$Datum<- format(as.Date(paste(regmatches(YukiDF$Omschrijving, regexpr("\\d{2}\\.\\d{2}",YukiDF$Omschrijving)),year,sep = "."), "%d.%m.%Y"),format = "%d-%m-%Y")
  matches <- regmatches(YukiDF$Omschrijving, gregexpr("\\d{2}\\.\\d{2}", YukiDF$Omschrijving))
  second_occurrence <- sapply(matches, function(x) ifelse(length(x) >= 2, x[2], NA))
  second_occurrence
  YukiDF$Rentedatum <-format(as.Date(paste(second_occurrence,year,sep = "."), "%d.%m.%Y"),format = "%d-%m-%Y")
  YukiDF$Omschrijving <-gsub("\\d{2}\\.\\d{2}\\s", "", YukiDF$Omschrijving) # now we have extracted the dates they can be removed
  pattern <- "\\b[0-9]{1,3}(?:\\.[0-9]{3})*(?:,[0-9]{2})\\b"
  YukiDF$Omschrijving<-gsub(pattern, "", YukiDF$Omschrijving) # also remove amounts from description
  YukiDF$Omschrijving <-gsub('\\"', "", YukiDF$Omschrijving)
  YukiDF$Bedrag[grep("PRLV SEPA REJ/IMP|VIRT SEPA", YukiDF$Omschrijving)]<--YukiDF$Bedrag[grep("PRLV SEPA REJ/IMP|VIRT SEPA", YukiDF$Omschrijving)]
  # parse names
  pattern <- "\\b[a-zA-Z]{5,}\\b" # first word of minimum 5 non digits
  matches <- regmatches(YukiDF$Omschrijving, regexpr(pattern, YukiDF$Omschrijving))
  YukiDF$Naam_tegenrekening[grep(pattern,YukiDF$Omschrijving)]<-matches
  matching_indices <- grep(pattern, YukiDF$Omschrijving)
  YukiDF$Naam_tegenrekening[matching_indices] <- matches[!is.na(matches)]
  matches
  #
  View(YukiDF)
  ColumnNames <-
    c(
      "IBAN",
      "Valuta",
      "Afschrift",
      "Datum",
      "Rentedatum",
      "Tegenrekening",
      "Naam tegenrekening",
      "Omschrijving",
      "Bedrag"
    )
  ofile
  write.table(
    YukiDF,
    file = ofile,
    quote = FALSE,
    sep = ";",
    dec = ".",
    row.names = FALSE,
    col.names = ColumnNames
  )
  stop("einde")
}

CreditCardDF <- as.data.frame(mCreditCard)
CreditCardDF <- cbind.data.frame(CreditCardDF, BedragDF)
View(mCreditCard)
# Delete empty Lines in CreditCard Dataframe 
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
CreditCardDF$Datum <- gsub("mrt", "mar", CreditCardDF$Datum, ignore.case = TRUE)
CreditCardDF$Datum <- gsub("mei", "may", CreditCardDF$Datum, ignore.case = TRUE)
CreditCardDF$Datum <- gsub("okt", "oct", CreditCardDF$Datum, ignore.case = TRUE)
CreditCardDF$Datum <- format(as.Date(CreditCardDF$Datum, "%d %b %Y"), "%d-%m-%Y")
CreditCardDF$Rentedatum <- gsub("mrt", "mar", CreditCardDF$Rentedatum, ignore.case = TRUE)
CreditCardDF$Rentedatum <- gsub("mei", "may", CreditCardDF$Rentedatum, ignore.case = TRUE)
CreditCardDF$Rentedatum <- gsub("okt", "oct", CreditCardDF$Rentedatum, ignore.case = TRUE)
CreditCardDF$Rentedatum <- format(as.Date(CreditCardDF$Rentedatum, "%d %b %Y"), "%d-%m-%Y")

# some magic to change the year of transaction that were in previous year compared to Date of statement since line do not have year indication
pyear<-as.character(as.numeric(year)-1)   # previous year to handle transaction around year end
nyear<-as.character(as.numeric(year)+1)   # next year to handle transaction around year end
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
# (?:^              # beginning of string
#     \d{1,3}       # one, two, or three digits
#   (?:
#       \.?         # optional separating period
#       \d{3}       # followed by exactly three digits
#   )*              # repeat this subpattern (.###) any number of times (including none at all)
#     (?:,\d{2})?   # optionally followed by a decimal comma and exactly two digits
#     $)            # End of string.
# |                 # ...or...
#   (?:^            # beginning of string
#      \d{1,3}      # one, two, or three digits
#    (?:
#        ,?         # optional separating comma
#        \d{3}      # followed by exactly three digits
#    )*             # repeat this subpattern (,###) any number of times (including none at all)
#      (?:\.\d{2})? # optionally followed by a decimal perioda and exactly two digits
#      $)           # End of string.
