appost <- function(){
  # pat <- utils::choose.dir(caption = "Seleziona la cartella dell'ordine")
  # setwd(pat)

  # Carica dati ordine ----
  cat("\014")
  #cat(rep("\n", 20))
  cat("

      ***************************
      *** BENVENUTI in AppOst ***
      ***************************


      Digitare il numero d'ordine e premere INVIO caricare il file 'Ordini.csv' scaricato da Teams


      ")
    # oppure digitare '0' (zero) per scaricare il file 'Elenco prodotti.xlsx'
  # (da compilare prima di generare RAS e lettera d'ordine)

  ordine <- readline()

  if(ordine==0){
    # pat <- utils::choose.dir()
    # setwd(pat)
    download.file("https://raw.githubusercontent.com/giovabubi/appost/main/models/Elenco%20prodotti.xlsx", destfile = "Elenco prodotti.xlsx", method = "curl")
    cat("\014")
    #cat(rep("\n", 20))
    cat("\014")
    cat("

    Documento 'Elenco prodotti.xlsx' generato e salvato in ", pat, "

    AppOst si chiuderà a breve. Grazie e arrivederci!

        ")
    Sys.sleep(5)
    quit(save="no")
  }

  if(file.exists("Ordini.csv")=="TRUE"){
    ordini <- read.csv("Ordini.csv", na.strings = "")
    pat <- getwd()
  }else{
    patfile <- utils::choose.files(default = "*.csv", caption = "Selezionare il file 'Ordini' scaricato da Teams")
    if(!require(stringr)) install.packages("stringr")
    n <- stringr::str_locate_all(patfile, "\\\\")
    m <- max(n[[1]])
    n <- paste0("(.{", m, "}).*")
    pat <- sub(n, "\\1", patfile)
    setwd(pat)
    ordini <- read.csv(patfile, na.strings = "")
  }

  if(!require(dplyr)) install.packages("dplyr")
  ordini <- dplyr::rename(ordini,
                          Prodotto=Descrizione.beni.servizi.lavori,
                          RDO=N..RDO.MePA,
                          sede=Sede)
  ordini$Fornitore..P.IVA <- as.character(ordini$Fornitore..P.IVA)
  ordini$CPV <- NULL
  ordini$CPV <- as.character(ordini$CPV..CPV)
  ordini$Importo.senza.IVA.num <- sub(",(..)$", "_\\1", ordini$Importo.senza.IVA)
  ordini$Importo.senza.IVA.num <- gsub("\\.", "", ordini$Importo.senza.IVA.num)
  ordini$Importo.senza.IVA.num <- gsub("_", ".", ordini$Importo.senza.IVA.num)
  ordini$Importo.senza.IVA.num <- as.numeric(ordini$Importo.senza.IVA.num)

  sc <- subset(ordini, ordini$Ordine.N.==ordine)

  # sc$Importo.senza.IVA.num <- sub(",(..)$", "_\\1", sc$Importo.senza.IVA)
  # sc$Importo.senza.IVA.num <- gsub("\\.", "", sc$Importo.senza.IVA.num)
  # sc$Importo.senza.IVA.num <- gsub("_", ".", sc$Importo.senza.IVA.num)
  # sc$Importo.senza.IVA.num <- as.numeric(sc$Importo.senza.IVA.num)
  sc$Aliquota.IVA.num <- as.numeric(ifelse(sc$Aliquota.IVA=='22%', 0.22,
                                           ifelse(sc$Aliquota.IVA=='10%', 0.1,
                                                  ifelse(sc$Aliquota.IVA=='4%', 0.04, 0))))
  sc$IVA <- sc$Importo.senza.IVA.num * sc$Aliquota.IVA.num
  sc$Importo.con.IVA <- sc$Importo.senza.IVA.num + sc$IVA
  sc$Importo.senza.IVA <- paste("€", format(sc$Importo.senza.IVA.num, format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
  sc$IVA <- paste("€", format(sc$IVA, format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
  sc$Importo.con.IVA <- paste("€", format(sc$Importo.con.IVA, format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))

  # Installa e carica pacchetti ----
  if(!require(officer)) install.packages("officer")
  if(!require(openxlsx)) install.packages("openxlsx")
  #if(!require(Microsoft365R)) install.packages("Microsoft365R")
  #if(!require(googledrive)) install.packages("googledrive")

  library(officer)
  library(openxlsx)
  library(dplyr)
  #library(Microsoft365R)
  #library(googledrive)

  ## Calcoli ----
  trattini <- "__________"
  sc[is.na(sc)] <- trattini
  attach(sc)

  fpt.b <- fp_text(bold = TRUE, font.family = 'Source Sans Pro')
  fpt.i <- fp_text(italic = TRUE, font.family = 'Source Sans Pro')
  fpt.bi <- fp_text(italic = TRUE, bold = TRUE, font.family = 'Source Sans Pro')
  PRODOTTO <- toupper(Prodotto)
  Prot..DaC <- tolower(Prot..DaC)

  if(sede=='BA'){
    sede1 <- 'Bari'
    sede2 <- 'Sede Secondaria di Bari'
    RSS <- 'Dott. Giovanni Nicola Bubici'
    RSS.email <- 'giovanninicola.bubici@cnr.it'
    RAMM <- 'Dott. Nicola Centorame'
    RAMM.email <- 'nicola.centorame@ipsp.cnr.it'
    al.RSS <- 'Al Responsabile della Sede Secondaria di Bari'
    firma.RSS <- 'Il Responsabile della Sede Secondaria di Bari'
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, via G. Amendola 122/D, 70126 Bari, Italia."
    nomina.RSS <- "177785 del 27/5/2024 e rettifica prot. 181568 del 29/5/2024 di nomina del dott. Giovanni Nicola Bubici quale Responsabile della Sede Secondaria di Bari dell’IPSP per il periodo dall’1/6/2024 al 31/12/2024"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Nicola Centorame quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
  }else if(sede=='TO'){
    sede1 <- 'Torino'
    sede2 <- 'Sede Secondaria di Torino'
    RSS <- 'Dott. Stefano Ghignone'
    RSS.email <- 'stefano.ghignone@cnr.it'
    RAMM <- "Dott.ssa Lucia Allione"
    RAMM.email <- 'lucia.allione@ipsp.cnr.it'
    al.RSS <- 'Al Responsabile della Sede Secondaria di Torino'
    firma.RSS <- 'Il Responsabile della Sede Secondaria di Torino'
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, viale Mattioli, 25, 10125 Torino, Italia."
    nomina.RSS <- "147145 del 3/5/2024 di nomina del dott. Stefano Ghignone quale Responsabile della Sede Secondaria di Torino dell’IPSP per il periodo dall’1/5/2024 al 31/12/2024"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della dott.ssa Lucia Allione quale Responsabile Amministrativo della Sede Secondaria di Torino dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
  }else if(sede=='NA'){
    sede1 <- 'Portici'
    sede2 <- 'Sede Secondaria di Portici'
    RSS <- 'Dott.ssa Michelina Ruocco'
    RSS.email <- 'michelina.ruocco@cnr.it'
    RAMM <- 'Dott. Ettore Magaldi'
    RAMM.email <- 'ettore.magaldi@ipsp.cnr.it'
    al.RSS <- 'Alla Responsabile della Sede Secondaria di Portici'
    firma.RSS <- 'La Responsabile della Sede Secondaria di Portici'
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, piazzale Enrico Fermi, 1, 80055 Portici (NA), Italia."
    nomina.RSS <- "147145 del 3/5/2024 di nomina della dott.ssa Michelina Ruocco quale Responsabile della Sede Secondaria di Portici dell’IPSP per il periodo dall’1/5/2024 al 31/12/2024"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Ettore Magaldi quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
  }else if(sede=='FI'){
    sede1 <- 'Sesto Fiorentino'
    sede2 <- 'Sede Secondaria di Sesto Fiorentino'
    RSS <- "Dott. Nicola Luchi"
    RSS.email <- "nicola.luchi@ipsp.cnr.it"
    RAMM <- "Sig.ra Francesca Pesciolini"
    RAMM.email <- 'francesca.pesciolini@ipsp.cnr.it'
    al.RSS <- 'Al Responsabile della Sede Secondaria di Sesto Fiorentino'
    firma.RSS <- 'Il Responsabile della Sede Secondaria di Sesto Fiorentino'
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, via Madonna del Piano, 10, 50019 Sesto F.no (FI), Italia."
    nomina.RSS <- "147145 del 3/5/2024 di nomina del dott. Nicola Luchi quale Responsabile della Sede Secondaria di Sesto Fiorentino dell’IPSP per il periodo dall’1/5/2024 al 31/12/2024"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della sig.ra Francesca Pesciolini quale Responsabile Amministrativo della Sede Secondaria di Sesto Fiorentino dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
  }else if(sede=='PD'){
    sede1 <- 'Legnaro'
    sede2 <- 'Sede Secondaria di Legnaro'
    RSS <- "Dott.ssa Laura Scarabel"
    RSS.email <- "laura.scarabel@ipsp.cnr.it"
    RAMM <- "Dott.ssa Lucia Allione"
    RAMM.email <- 'lucia.allione@ipsp.cnr.it'
    al.RSS <- 'Al Responsabile della Sede Secondaria di Legnaro'
    firma.RSS <- 'Il Responsabile della Sede Secondaria di Legnaro'
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, viale dell’Università, 16, 35020 Legnaro (PD), Italia."
    nomina.RSS <- "147145 del 3/5/2024 di nomina della dott.ssa Laura Scarabel quale Responsabile della Sede Secondaria di Legnaro dell’IPSP per il periodo dall’1/5/2024 al 31/12/2024"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della dott.ssa Lucia Allione quale Responsabile Amministrativo della Sede Secondaria di Legnaro dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
  }else if(sede=='TOsi'){
    sede1 <- 'Torino'
    sede2 <- 'Sede Istituzionale'
    RSS <- 'Dott. Francesco Di Serio'
    RSS.email <- 'francesco.diserio@cnr.it'
    RAMM <- 'Dott. Josè Saporita'
    RAMM.email <- 'jose.saporita@ipsp.cnr.it'
    al.RSS <- "Al Direttore dell'IPSP"
    firma.RSS <- "Il Direttore"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, Strada delle Cacce, 73, 10135 Torino, Italia."
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Josè Saporita quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
  }

  if(Scelta.fornitore=='Avviso pubblico'){
    preventivo.individuato <- paste0("stato individuato per motivazioni tecnico-scientifiche e di economicità tra i preventivi di ",
                                     Prot..preventivi.avviso,
                                     " pervenuti in seguito all'avviso pubblico prot. ",
                                     Prot..avviso.pubblico,
                                     ";")
  }else if(Scelta.fornitore=='Più preventivi'){
    preventivo.individuato <- "stato individuato a seguito di indagine informale di mercato effettuata su MePA, mercato libero e/o cataloghi accessibili in rete con esito allegato alla richiesta medesima;"
  }else{
    preventivo.individuato <- "allegato alla richiesta medesima;"
  }

  if(Rotazione.fornitore=="Importo <5.000€"){
    rotazione.individuata <- "che, in relazione a quanto indicato all'art. 49, comma 6, del Codice è possibile derogare dall'applicazione del principio di rotazione in caso di affidamenti di importo inferiore a euro 5.000,00;"
  }else if(Rotazione.fornitore=="Avviso pubblico"){
    rotazione.individuata <- "che non si applica il principio di rotazione in quanto è stata espletata un'indagine di mercato aperta alla partecipazione di tutti gli operatori economici in possesso di tutti i requisiti richiesti;"
  }else if(Rotazione.fornitore=="Non è il contraente uscente"){
    rotazione.individuata <- "che in applicazione del principio di rotazione l'operatore economico individuato non è il contraente uscente;"
  }else{
    rotazione.individuata <- "che è possibile procedere all’affidamento al contraente uscente poiché non trova applicazione il principio di rotazione in conseguenza della particolare struttura del mercato e dell'effettiva assenza di alternative e che l'affidatario medesimo ha svolto;"
  }

  if(Motivo.fuori.MePA==">5.000€ beni ICT"){
    ICT.testo <- " che trattasi di beni funzionalmente destinati all’attività di ricerca e che pertanto trovano applicazione le disposizioni di cui all’art. 4 comma 1 lettera b) del D.L. 126/2019 convertito in L. 159/2019;"
  }else if(Motivo.fuori.MePA==">5.000€ beni nonICT"){
    ICT.testo <- " che trattasi di beni funzionalmente destinati all’attività di ricerca e che pertanto trovano applicazione le disposizioni di cui all’art. 10 comma 3 del D.lgs. 218/2016;"
  }else{
    ICT.testo <- ""
  }

  if(Tipo.acquisizione=='Beni'){
    bene <- 'bene'
    beni <- 'beni'
    della.fornitura <- 'della fornitura'
    la.fornitura <- 'la fornitura'
    fornitura.consegnata <- 'la fornitura dovrà essere consegnata'
    materiale.conforme <- "che il materiale è conforme all’ordine e perfettamente funzionante e utilizzabile."
  }else if(Tipo.acquisizione=='Servizi'){
    bene <- 'servizio'
    beni <- 'servizi'
    della.fornitura <- 'del servizio'
    la.fornitura <- 'il servizio'
    fornitura.consegnata <- 'il servizio dovrà essere prestato'
    materiale.conforme <- "che il servizio è conforme all’ordine e completamente prestato."
  }else if(Tipo.acquisizione=='Lavori'){
    bene <- 'lavoro'
    beni <- 'lavori'
    della.fornitura <- 'del lavoro'
    la.fornitura <- 'il lavoro'
    fornitura.consegnata <- 'il lavoro dovrà essere svolto'
    materiale.conforme <- "che il lavoro è conforme all’ordine e completamente svolto."
  }

  if(Richiedente..Sesso=='M'){
    Dott.ric <- 'Dott.'
    dott.ric <- 'dott.'
    il.dott.ric <- 'il dott.'
    al.ric <- 'al'
    dal.ric <- 'dal dott.'
    Nato.ric <- 'Nato a'
    nato.ric <- 'nato a'
    assegna <- 'assegnatario'
    sottoscritto.ric <- 'Il sottoscritto'
  }else{
    Dott.ric <- 'Dott.ssa'
    dott.ric <- 'dott.ssa'
    il.dott.ric <- 'la dott.ssa'
    al.ric <- 'alla'
    dal.ric <- 'dalla dott.ssa'
    Nato.ric <- 'Nata a'
    nato.ric <- 'nata a'
    assegna <- 'assegnataria'
    sottoscritto.ric <- 'La sottoscritta'
  }
  if(Responsabile.progetto..Sesso=='M'){
    Dott.resp <- "Dott."
    dott.resp <- 'dott.'
    il.dott.resp <- 'il dott.'
    al.resp <- 'al'
    dal.resp <- 'dal dott.'
    Nato.resp <- 'Nato a'
    nato.resp <- 'nato a'
    sottoscritto.resp <- 'Il sottoscritto'
  }else{
    Dott.resp <- "Dott.ssa"
    dott.resp <- 'dott.ssa'
    il.dott.resp <- 'la dott.ssa'
    al.resp <- 'alla'
    dal.resp <- 'dalla dott.ssa'
    Nato.resp <- 'Nata a'
    nato.resp <- 'nata a'
    sottoscritto.resp <- 'La sottoscritta'
  }
  if(RUP..Sesso=='M'){
    Dott.rup <- 'Dott.'
    dott.rup <- 'dott.'
    il.dott.rup <- 'il dott.'
    al.rup <- 'al'
    dal.rup <- 'dal dott.'
    Nato.rup <- 'Nato a'
    nato.rup <- 'nato a'
    sottoscritto.rup <- 'Il sottoscritto'
    nominato <- "stato nominato"
  }else{
    Dott.rup <- 'dott.ssa'
    dott.rup <- 'dott.ssa'
    il.dott.rup <- 'la dott.ssa'
    al.rup <- 'alla'
    dal.rup <- 'dalla dott.ssa'
    Nato.rup <- 'Nata a'
    nato.rup <- 'nata a'
    sottoscritto.rup <- 'La sottoscritta'
    nominato <- "stata nominata"
  }

  da <- as.character(Sys.Date())
  y <- sub("(....)-(..)-(..)",  "/\\1", da)
  y2 <- sub("(....)-(..)-(..)",  "\\1", da)
  da <- sub("(....)-(..)-(..)",  "\\3/\\2/\\1", da)

  if(CUP!=trattini){
    CUP1 <- paste0(" (CUP ", CUP, ")")
    Progetto1 <- paste0(Progetto, " (CUP ", CUP, ")")
    CUP2 <- CUP
  }else{
    CUP1 <- ""
    Progetto1 <- Progetto
    CUP2 <- ""
  }

  if(CUI!=trattini){
    CUI1 <- paste0(", CUI ", CUI)
    CUI2 <- CUI
    CUI3 <- CUI
  }else{
    CUI1 <- ""
    CUI2 <- ""
    CUI3 <- trattini
  }

  if(Tipo.ordine=='Ordine diretto MePA'){
    ordine.trattativa.scelta <- paste0(", ordine diretto MePA N° ", RDO)
    ordine.trattativa.scelta2 <- "Ordine diretto MePA N°"
    ordine.trattativa.scelta.ldo1 <- "Ordine diretto MePA N° "
    ordine.trattativa.scelta.ldo2 <- RDO
    ordine.trattativa.scelta.pres <- paste0("l'ordine diretto MePA N° ", RDO, ";")
  }else if(Tipo.ordine=='Trattativa diretta MePA'){
    ordine.trattativa.scelta <- paste0(", trattativa diretta MePA N° ", RDO)
    ordine.trattativa.scelta2 <- "Trattativa diretta MePA N°"
    ordine.trattativa.scelta.ldo1 <- "Trattativa diretta MePA N° "
    ordine.trattativa.scelta.ldo2 <- RDO
    ordine.trattativa.scelta.pres <- paste0("la trattativa diretta MePA N° ", RDO, ";")
  }else if(Tipo.ordine=='RDO MePA'){
    ordine.trattativa.scelta <- paste0(", RDO MePA N° ", RDO)
    ordine.trattativa.scelta2 <- "RDO MePA N°"
    ordine.trattativa.scelta.ldo1 <- "RDO MePA N° "
    ordine.trattativa.scelta.ldo2 <- RDO
    ordine.trattativa.scelta.pres <- paste0("la RDO MePA N° ", RDO, ";")
  }else{
    ordine.trattativa.scelta <- ""
    ordine.trattativa.scelta2 <- ""
    ordine.trattativa.scelta.ldo1 <- "Vs. offerta "
    ordine.trattativa.scelta.ldo2 <- Preventivo.fornitore
    ordine.trattativa.scelta.pres <- paste0("l'offerta ", Preventivo.fornitore, " dell'operatore economico ", Fornitore, ", P.I/C.F. ", Fornitore..P.IVA, ";")
  }

  if(Richiedente==Responsabile.progetto){firma.RAS <- "(Richiedente l’ordine, responsabile del progetto e titolare dei fondi)"}else{firma.RAS <- "(Richiedente l’ordine)"}

  CAMPO.OE <- paste0(Fornitore, ', P.I/C.F. ', Fornitore..P.IVA, ', con sede in ', Fornitore..Sede, ', tel. ', Fornitore..Telefono, ', PEC ', Fornitore..PEC, ', e-mail ', Fornitore..E.mail)
  CAMPO.OE1 <- trattini
  CAMPO.OE2 <- trattini
  all.OE <- paste0(", all'operatore economico ", Fornitore, " (P.IVA ", Fornitore..P.IVA, ")")
  if(Scelta.fornitore=='Singolo preventivo'){
    CAMPO.OE1 <- trattini
    CAMPO.OE2 <- CAMPO.OE
  }else if(Scelta.fornitore=='Più preventivi'){
    CAMPO.OE1 <- CAMPO.OE
    CAMPO.OE2 <- trattini
  }else if(Scelta.fornitore=='Avviso pubblico'){
    all.OE <- ''
  }else{
    all.OE <- ''
  }

  int.doc <- toupper(paste0("Affidamento diretto, ai sensi dell’art. 50 del D.Lgs. N. 36/2023, ",
                    della.fornitura, " di “", Prodotto, "” (CIG ", CIG, CUI1, ", ", Pagina.web, ") ",
                    "nell'ambito del progetto “", Progetto, "”", CUP1,
                    ordine.trattativa.scelta,
                    ", ordine CNR-IPSP-", sede, " N° ", ordine, y, ".", sep=""))

  pre.nome.file <- paste0("Ordine CNR-IPSP-", sede, " ", ordine, "_", y2, " - ")

  # Ultimi DocOE ----
  ultimi <- subset(ordini, ordini$Fornitore==sc$Fornitore)
  ultimi <- dplyr::select(ultimi, Ordine.N., Fornitore, Prot..DocOE)
  ultimi$Prot..DocOE[which(ultimi$Prot..DocOE=="")] <- NA
  ultimi$data <- sub(".* ([0-9])", "\\1", ultimi$Prot..DocOE)
  ultimi$data <- as.POSIXct(ultimi$data, tz="CET", format = "%d/%m/%Y")
  today <- format(Sys.Date(), "%d/%m/%Y", tz="CET")
  today <- as.POSIXct(today, tz="CET", format = "%d/%m/%Y")
  ultimi$diff <- as.numeric(round(today - ultimi$data, 0))
  ultimi$diff[which(is.na(ultimi$data))] <- 999
  ultimi <- subset(ultimi, ultimi$Ordine.N.!=ordine)
  lng.doc <- length(ultimi$Fornitore)
  ultimi <- ultimi[order(ultimi$diff),]
  ultimi.ordine <- ultimi$Ordine.N.[1]
  ultimi.prot <- ultimi$Prot..DocOE[1]
  ultimi.recente <- ultimi$diff[1]
  if(lng.doc==0){ultimi.recente <- 999}

  # Rotazione fornitore ----
  if(sc$Importo.senza.IVA.num<=5000){
    ordini.fascia <- subset(ordini, ordini$Importo.senza.IVA.num<=5000)
    }else if(sc$Importo.senza.IVA.num>5000 & sc$Importo.senza.IVA.num<=40000){
      ordini.fascia <- subset(ordini, ordini$Importo.senza.IVA.num>5000 & ordini$Importo.senza.IVA.num<=40000)
    }else if(sc$Importo.senza.IVA.num>40000){
      ordini.fascia <- subset(ordini, ordini$Importo.senza.IVA.num<=5000)
    }

  rota <- subset(ordini.fascia, ordini.fascia$CPV==sc$CPV)
  rota <- dplyr::select(rota, Ordine.N., Data, Fornitore, CPV, Prodotto, Importo.senza.IVA)
  rota$Data <- as.POSIXct(rota$Data, tz="CET", format = "%d/%m/%Y")
  data.ordine <- subset(rota, rota$Ordine.N.==ordine)
  data.ordine <- data.ordine$Data
  rota <- subset(rota, rota$Data<=data.ordine)
  rota$CPV.iniz <- sub("(...).*", "\\1", rota$CPV)
  rota <- subset(rota, rota$Ordine.N.!=ordine)
  rota <- rota[order(rota$Data),]
  lng <- length(rota$Ordine.N.)
  rota <- rota[lng,]
  rota.display <- dplyr::select(rota, Ordine.N., Fornitore, CPV, Prodotto, Importo.senza.IVA)
  ordine.uscente <- rota$Ordine.N.
  fornitore.uscente <- rota$Fornitore
  cpv.usente <- rota$CPV.iniz
  prodotto.uscente <- rota$Prodotto
  importo.uscente <- rota$Importo.senza.IVA
  if(lng==0){fornitore.uscente <- "nessuno"}

  # Scarica Modello.docx da GoogleDrive ---
  # drive_deauth()
  # drive_user()
  # modello <- drive_get(as_id("1AOrViONf-0tZI22Hzn1dCNDcn_xxPag-"))
  # drive_download(modello, overwrite = TRUE)
  # doc.ras <- read_docx(modello$name)
  # doc.avv <- read_docx(modello$name)
  # doc.all <- read_docx(modello$name)
  # doc.dac <- read_docx(modello$name)
  # doc.prov.imp <- read_docx(modello$name)
  # doc.pag <- read_docx(modello$name)
  # doc.pi <- read_docx(modello$name)
  # doc.cc <- read_docx(modello$name)
  # doc.dgue <- read_docx(modello$name)
  # doc.dpcm <- read_docx(modello$name)
  # doc.doh <- read_docx(modello$name)
  # doc.com.cig <- read_docx(modello$name)
  # doc.ai <- read_docx(modello$name)
  # doc.ldo <- read_docx(modello$name)
  # doc.dic.pres <- read_docx(modello$name)
  # doc.prov.liq <- read_docx(modello$name)
  # file.remove(modello$name)

  # Scarica Modello.docx da Github ----
  download.file("https://raw.githubusercontent.com/giovabubi/appost/main/models/Modello.docx", destfile = "Modello.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
  doc.ras <- read_docx("Modello.docx")
  doc.avv <- read_docx("Modello.docx")
  doc.all <- read_docx("Modello.docx")
  doc.dac <- read_docx("Modello.docx")
  doc.prov.imp <- read_docx("Modello.docx")
  doc.pag <- read_docx("Modello.docx")
  doc.pi <- read_docx("Modello.docx")
  doc.cc <- read_docx("Modello.docx")
  doc.part.qual <- read_docx("Modello.docx")
  doc.qual <- read_docx("Modello.docx")
  doc.aus <- read_docx("Modello.docx")
  doc.dpcm <- read_docx("Modello.docx")
  doc.doh <- read_docx("Modello.docx")
  doc.bollo <- read_docx("Modello.docx")
  doc.com.cig <- read_docx("Modello.docx")
  doc.ai <- read_docx("Modello.docx")
  doc.ldo <- read_docx("Modello.docx")
  doc.dic.pres <- read_docx("Modello.docx")
  doc.prov.liq <- read_docx("Modello.docx")
  file.remove("Modello.docx")

  # Genera RAS ----
  ras <- function(){
    cat("\014")
    #cat(rep("\n", 20))
    cat("\014")

    if(Fornitore==fornitore.uscente){
      cat(paste0(
        "***** ATTENZIONE *****\n",
        Fornitore, " è il fornitore uscente.\n",
        "L'ultimo ordine (n° ", ordine.uscente, ") per questa categoria merceologica (prime tre cifre del CPV: ", cpv.usente, ") è stato affidato a questo operatore economico per l'acquisto di '", prodotto.uscente, "' e un importo di € ", importo.uscente, ".\n"))
      if(Rotazione.fornitore=="Non è il contraente uscente"){
        cat("In FluOr è stato erroneamente indicato 'Non è il contraente uscente'. Si prega di apportare la dovuta correzione.\n")
      }else if(Rotazione.fornitore=="Particolare struttura del mercato"){
        cat("L'ordine può procedere poichè è stato indicato 'Particolare struttura del mercato'.\n")
      }else if(Rotazione.fornitore=="Importo <5.000€" & Importo.senza.IVA.num<5000){
        cat("L'ordine può procedere poichè è stata specificata la deroga alla rotazione dei fornitori per ordini <5.000 €.\n")
      }else if(Rotazione.fornitore=="Importo <5.000€" & Importo.senza.IVA.num>=5000){
        cat("E' stata specificata la deroga alla rotazione dei fornitori per ordini <5.000 €, ma l'ordine è superiore a questo importo. Si prega di apportare la dovuta correzione.\n")
      }
        cat("*********************\n",
        " Premere INVIO per proseguire")
      readline()
    }

    if(file.exists("Elenco prodotti.xlsx")=="FALSE"){
    cat("

    Premere INVIO per caricare il file Excel con l'elenco dei prodotti

      ")
    inpt <- readline()
    pr <- read.xlsx(utils::choose.files(default = "*.xlsx"))
    }else{
      pr <- read.xlsx("Elenco prodotti.xlsx")
    }

    pr <- pr[,1:5]
    colnames(pr) <- c("Quantità", "Descrizione", "Costo unitario senza IVA", "Importo senza IVA", "Inv./Cons.")
    pr <- subset(pr, !is.na(pr$Quantità))
    pr$`Inv./Cons.`[which(is.na(pr$`Inv./Cons.`))] <- ""
    pr$`Costo unitario senza IVA` <- paste("€", format(as.numeric(pr$`Costo unitario senza IVA`), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    pr$`Importo senza IVA` <- paste("€", format(as.numeric(pr$`Importo senza IVA`), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))

    prt <- pr[,-5]
    colnames(prt) <- c("Quantità", "Descrizione", "Costo unitario", "Importo")

    doc <- doc.ras |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_reach("CAMPO.DEST.RAS.SEDE") |>
      body_replace_all_text("CAMPO.DEST.RAS.SEDE", al.RSS, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DEST.RAS.RSS") |>
      body_replace_all_text("CAMPO.DEST.RAS.RSS", RSS, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DEST.RAS.EMAIL") |>
      body_replace_all_text("CAMPO.DEST.RAS.EMAIL", RSS.email, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DELLA.FORNITURA") |>
      body_replace_all_text("CAMPO.DELLA.FORNITURA", della.fornitura, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.PRODOTTO") |>
      body_replace_all_text("CAMPO.PRODOTTO", Prodotto, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.PROGETTO") |>
      body_replace_all_text("CAMPO.PROGETTO", Progetto1, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.SOTTOSCRITTO") |>
      body_replace_all_text("CAMPO.SOTTOSCRITTO", sottoscritto.ric, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.RICHIEDENTE") |>
      body_replace_all_text("CAMPO.RICHIEDENTE", Richiedente, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.BENI") |>
      body_replace_all_text("CAMPO.BENI", beni, only_at_cursor = FALSE) |>
      body_add_par("") |>
      body_add_table(pr, style = "Stile1") |>
      cursor_reach("CAMPO.SEDE") |>
      body_replace_all_text("CAMPO.SEDE", sede1, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.IMPORTO.SENZA.IVA") |>
      body_replace_all_text("CAMPO.IMPORTO.SENZA.IVA", Importo.senza.IVA, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.VOCE") |>
      body_replace_all_text("CAMPO.VOCE", Voce.di.spesa, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.PROGETTO") |>
      body_replace_all_text("CAMPO.PROGETTO", Progetto, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.GAE") |>
      body_replace_all_text("CAMPO.GAE", GAE, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.CUP") |>
      body_replace_all_text("CAMPO.CUP", CUP2, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.OE1") |>
      body_replace_all_text("CAMPO.OE1", CAMPO.OE1, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.OE2") |>
      body_replace_all_text("CAMPO.OE2", CAMPO.OE2, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DATA") |>
      body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da)), pos = "on") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(Dott.ric), ftext(" "), ftext(Richiedente)), style = "Firma 2") |>
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2") |>
      cursor_reach("CAMPO.DATA") |>
      body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da)), pos = "on") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(Dott.ric), ftext(" "), ftext(Richiedente)), style = "Firma 2") |>
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2") |>
      cursor_reach("CAMPO.LA.FORNITURA") |>
      body_replace_all_text("CAMPO.LA.FORNITURA", la.fornitura, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.FORNITORE") |>
      body_replace_all_text("CAMPO.FORNITORE", Fornitore, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.IMPORTO") |>
      body_replace_all_text("CAMPO.IMPORTO", Importo.senza.IVA, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DELLA.FORNITURA") |>
      body_replace_all_text("CAMPO.DELLA.FORNITURA", della.fornitura, only_at_cursor = TRUE)

    ## Dich. Ass. Rich. ----
    doc <- cursor_reach(doc, "SEZIONE.DICH.ASS.RICH.")
    doc <- doc |>
      body_add_fpar(fpar(ftext("All’"),
      ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50 DEL D.LGS. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" DI “"),
                         ftext(PRODOTTO),
                         ftext("”"),
                         ftext(", nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", ORDINE CNR-IPSP-"),
                         ftext(sede),
                         ftext(" N° "),
                         ftext(ordine),
                         ftext(y)), style = "Maiuscolo") |>
      body_add_fpar(fpar(ftext("AUTODICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.ric), ftext(" "), ftext(Richiedente, fpt.b), ftext(", "),
                         ftext(nato.ric), ftext(" "), ftext(Richiedente..Luogo.di.nascita), ftext(", il "),
                         ftext(Richiedente..Data.di.nascita), ftext(", codice fiscale "), ftext(Richiedente..Codice.fiscale), ftext(",")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di richiedente la fornitura di “"),
                         ftext(Prodotto),
                         ftext("”, ordine CNR-IPSP-"),
                         ftext(sede),
                         ftext(" N° "),
                         ftext(ordine),
                         ftext(y),
                         ftext(all.OE),
                         ftext(", nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                         ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di non trovarsi, rispetto al ruolo ricoperto ed alle funzioni svolte, in alcuna delle situazioni di conflitto di interessi, anche potenziale, di cui all’art. 16 del D.lgs. n. 36/2023, né nelle ipotesi previste dall’art. 35-bis, del D.lgs. n. 165/2001, tali da ledere l’imparzialità e l’immagine dell’agire dell’amministrazione;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("di impegnarsi a comunicare qualsiasi conflitto d’interesse che possa insorgere durante il presente affidamento o nella fase esecutiva del contratto;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("di impegnarsi ad astenersi prontamente dalla prosecuzione dell’affidamento diretto nel caso emerga un conflitto d’interesse;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("DICHIARA ALTRESÌ")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di aver preso piena cognizione del D.P.R. 16 aprile 2013, n. 62 e delle norme in esso contenute, nonché del Codice di comportamento dei dipendenti del Consiglio Nazionale delle Ricerche adottato con delibera del Consiglio di Amministrazione n° 137/2017;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("SI IMPEGNA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("a non utilizzare a fini privati le informazioni di cui dispone in ragione del ruolo ricoperto, a non divulgarle al di fuori dei casi consentiti e ad evitare situazioni e comportamenti che possano ostacolare il corretto adempimento della funzione sopra descritta;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("a comunicare tempestivamente eventuali variazioni del contenuto della presente dichiarazione e a rendere, se del caso, una nuova dichiarazione sostitutiva.")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("La presente dichiarazione è resa ai sensi e per gli effetti dell’art. 6-bis Legge 241/1990, degli artt. 6 e 7 del D.P.R. 16 aprile 2013, n. 62, dell’art. 53, comma 14, del D. Lgs. n° 165/2001, dell’art. 15, comma 1, lettera c) del D. Lgs. n° 33/2013 e dell’art. 20 del D. Lgs. n° 39/2013.")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sede1), ftext(", "),ftext(da)), style = "Normal") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(paste0(Dott.ric," ", Richiedente), run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2")

    ## Dich. Ass. Resp. ----
    if(Richiedente!=Responsabile.progetto){
      doc <- doc |>
        cursor_reach("SEZIONE.DICH.ASS.RESP.") |>
        body_add_fpar(fpar(ftext("All’"),
                           ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
        body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50 DEL D.LGS. N. 36/2023, "),
                           ftext(della.fornitura), ftext(" DI “"),
                           ftext(PRODOTTO),
                           ftext("”"),
                           ftext(", nell'ambito del progetto “"),
                           ftext(Progetto),
                           ftext("”"),
                           ftext(CUP1),
                           ftext(", ORDINE CNR-IPSP-"),
                           ftext(sede),
                           ftext(" N° "),
                           ftext(ordine),
                           ftext(y)), style = "Maiuscolo") |>
        body_add_fpar(fpar(ftext("AUTODICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext(sottoscritto.resp), ftext(" "), ftext(Responsabile.progetto, fpt.b), ftext(", "),
                           ftext(nato.resp), ftext(" "), ftext(Responsabile.progetto..Luogo.di.nascita), ftext(", il "),
                           ftext(Responsabile.progetto..Data.di.nascita), ftext(", codice fiscale "), ftext(Responsabile.progetto..Codice.fiscale), ftext(",")), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b),
                           ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di titolare dei fondi e responsabile del progetto di ricerca “"),
                           ftext(Progetto), ftext("”"), ftext(CUP1),
                           ftext(", in relazione alla fornitura di “"),
                           ftext(Prodotto),
                           ftext("”, ordine CNR-IPSP-"),
                           ftext(sede),
                           ftext(" N° "),
                           ftext(ordine),
                           ftext(y),
                           ftext(all.OE),
                           ftext(", consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                           ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("di non trovarsi, rispetto al ruolo ricoperto ed alle funzioni svolte, in alcuna delle situazioni di conflitto di interessi, anche potenziale, di cui all’art. 16 del D.lgs. n. 36/2023, né nelle ipotesi previste dall’art. 35-bis, del D.lgs. n. 165/2001, tali da ledere l’imparzialità e l’immagine dell’agire dell’amministrazione;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("di impegnarsi a comunicare qualsiasi conflitto d’interesse che possa insorgere durante il presente affidamento o nella fase esecutiva del contratto;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("di impegnarsi ad astenersi prontamente dalla prosecuzione dell’affidamento diretto nel caso emerga un conflitto d’interesse;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("DICHIARA ALTRESÌ")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("di aver preso piena cognizione del D.P.R. 16 aprile 2013, n. 62 e delle norme in esso contenute, nonché del Codice di comportamento dei dipendenti del Consiglio Nazionale delle Ricerche adottato con delibera del Consiglio di Amministrazione n° 137/2017;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("SI IMPEGNA")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("a non utilizzare a fini privati le informazioni di cui dispone in ragione del ruolo ricoperto, a non divulgarle al di fuori dei casi consentiti e ad evitare situazioni e comportamenti che possano ostacolare il corretto adempimento della funzione sopra descritta;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("a comunicare tempestivamente eventuali variazioni del contenuto della presente dichiarazione e a rendere, se del caso, una nuova dichiarazione sostitutiva.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("La presente dichiarazione è resa ai sensi e per gli effetti dell’art. 6-bis Legge 241/1990, degli artt. 6 e 7 del D.P.R. 16 aprile 2013, n. 62, dell’art. 53, comma 14, del D. Lgs. n° 165/2001, dell’art. 15, comma 1, lettera c) del D. Lgs. n° 33/2013 e dell’art. 20 del D. Lgs. n° 39/2013.")), style = "Normal") |>
        body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da)), style = "Normal") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(paste0(Dott.resp," ",Responsabile.progetto), run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(Responsabile del progetto e titolare dei fondi)")), style = "Firma 2")

      b <- cursor_reach(doc, "CAMPO.CUP.LDO")
      b <- b$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which
      doc <- cursor_reach(doc, "CAMPO.CUP.LDO")
      for(i in 1:(e-b)){
        doc <- body_remove(doc)
      }
      doc <- cursor_end(doc)
      doc <- cursor_backward(doc)
      doc <- cursor_backward(doc)
      doc <- cursor_backward(doc)
      doc <- body_remove(doc)
    }else{
      b <- cursor_reach(doc, "SEZIONE.DICH.ASS.RESP.")
      b <- b$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which
      doc <- cursor_reach(doc, "SEZIONE.DICH.ASS.RESP.")
      for(i in 1:(e-b)){
        doc <- body_remove(doc)
      }
      doc <- cursor_end(doc)
      doc <- cursor_backward(doc)
      doc <- cursor_backward(doc)
      doc <- cursor_backward(doc)
      doc <- body_remove(doc)
      doc <- body_remove(doc)
    }
    print(doc, target = paste0(pre.nome.file, "1 RAS.docx"))

    cat("\014")
    #cat(rep("\n", 20))
    cat("\014")
    cat("

    Documento '", pre.nome.file, "1 RAS.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, GAE, Richiedente..Luogo.di.nascita,
                           Richiedente..Codice.fiscale, Responsabile.progetto, Responsabile.progetto..Luogo.di.nascita, Responsabile.progetto..Codice.fiscale)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
    }

    ## Avviso pubblico ----
    if(Scelta.fornitore=='Avviso pubblico'){
      doc <- doc.avv |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.DEST.RAS.SEDE") |>
        body_add_fpar(fpar(ftext("AVVISO")), style = "heading 1", pos = "on") |>
        body_add_fpar(fpar(ftext("Indagine esplorativa di mercato volta a raccogliere preventivi finalizzati all’affidamento "),
                           ftext(della.fornitura),
                           ftext(" di “"),
                           ftext(Prodotto),
                           ftext("” nell'ambito del progetto “"),
                           ftext(Progetto1),
                           ftext("”.")), style = "Titolo avviso") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Premesse e finalità")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("La Stazione Appaltante ISTITUTO PER LA PROTEZIONE SOSTENIBILE DELLE PIANTE del CNR intende procedere, a mezzo della presente indagine esplorativa, all’individuazione di un operatore economico a cui affidare eventualmente il servizio di cui all’oggetto, ai sensi dell’art. 50, comma 1 del d.lgs. 36/2023.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Il presente avviso, predisposto nel rispetto dei principi di libera concorrenza, non discriminazione, trasparenza, proporzionalità e pubblicità, non costituisce invito a partecipare a gara pubblica, né un’offerta al pubblico (art. 1336 del Codice civile) o promessa al pubblico (art. 1989 del Codice civile), ma ha lo scopo di esplorare le possibilità offerte dal mercato al fine di affidare direttamente il servizio.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("L’indagine in oggetto non comporta l’instaurazione di posizioni giuridiche ovvero obblighi negoziali. Il presente avviso, pertanto, non vincola in alcun modo questa Stazione Appaltante che si riserva, comunque, la facoltà di sospenderlo, modificarlo o annullarlo e di non dar seguito al successivo affidamento, senza che gli operatori economici possano vantare alcuna pretesa.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("I preventivi ricevuti si intenderanno impegnativi per gli operatori economici per un periodo di massimo di 60 giorni naturali e consecutivi, mentre non saranno in alcun modo impegnativi per la Stazione Appaltante, per la quale resta salva la facoltà di procedere o meno a successive e ulteriori richieste di preventivi volte all’affidamento del servizio di cui all’oggetto.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Oggetto "), ftext(della.fornitura)), style = "heading 3") |>
        body_add_fpar(fpar(ftext("L’oggetto "), ftext(della.fornitura), ftext(" è _____.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("La consegna dovrà avvenire presso _____ entro _____.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("[Specificare tutte le caratteristiche del bene/servizio/lavoro, nonchè modalità e tempi di consegna, così che gli operatori economici possano presentare offerte comparabili e la stazione appaltante possa scegliere il preventivo più adatto in base ai criteri richiesti in fase di avviso pubblico]", fpt.i)), style = "Normal") |>
        body_add_fpar(fpar(ftext("Requisiti")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Possono inviare il proprio preventivo gli operatori economici in possesso di:")), style = "Normal") |>
        body_add_fpar(fpar(ftext("requisiti di ordine generale di cui al Capo II, Titolo IV del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("requisiti d’idoneità professionale come specificato all’art. 100, comma 3 del D.lgs. n. 36/2023: iscrizione nel registro della camera di commercio, industria, artigianato e agricoltura o nel registro delle commissioni provinciali per l’artigianato o presso i competenti ordini professionali per un’attività pertinente anche se non coincidente con l’oggetto dell’appalto. All’operatore economico di altro Stato membro non residente in Italia è richiesto di dichiarare ai sensi del testo unico delle disposizioni legislative e regolamentari in materia di documentazione amministrativa, di cui al decreto del Presidente della Repubblica del 28 dicembre 2000, n. 445;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("pregresse e documentate esperienze analoghe anche se non coincidenti con quelle oggetto dell’appalto.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Valore dell'affidamento")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("La Stazione Appaltante ha stimato per l’affidamento di cui all’oggetto un importo massimo pari a "),
                           ftext(Importo.senza.IVA), ftext(" oltre IVA.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Modalità di presentazione del preventivo")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Gli operatori economici in possesso dei requisiti sopra indicati potranno inviare il proprio preventivo, corredato della dichiarazione attestante il possesso dei requisiti predisposta secondo il modello allegato al presente avviso (allegato 1), entro e non oltre 15 giorni dalla pubblicazione del presente avviso a mezzo PEC all’indirizzo protocollo.ipsp@pec.cnr.it e per conoscenza a "),
                           ftext(Richiedente..E.mail), ftext(" e jose.saporita@ipsp.cnr.it indicando nell’oggetto “Att.ne "),
                           ftext(dott.ric), ftext(" "), ftext(Richiedente),
                           ftext(": preventivo relativo all’avviso pubblico per "),
                           ftext(la.fornitura), ftext(" di "), ftext(Prodotto), ftext("”.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("La documentazione trasmessa dovrà essere sottoscritta digitalmente con firma qualificata da un legale rappresentante/procuratore in grado di impegnare l’operatore economico.")), style = "Normal")

      if(sede!='TOsi'){
        doc <- doc |>
        body_add_fpar(fpar(ftext("Gli operatori economici stranieri non residenti in Italia, sprovvisti di posta elettronica certificata, possono spedire il preventivo e la dichiarazione in lingua italiana all’indirizzo "),
                           ftext(RAMM.email), ftext(" e per conoscenza a "),
                           ftext(Richiedente..E.mail), ftext(" e jose.saporita@ipsp.cnr.it indicando nell’oggetto “Att.ne "), ftext(dott.ric), ftext(" "), ftext(Richiedente), ftext(": preventivo relativo all’avviso pubblico per "), ftext(la.fornitura), ftext(" di "), ftext(Prodotto), ftext("”.")), style = "Normal")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("Gli operatori economici stranieri non residenti in Italia, sprovvisti di posta elettronica certificata, possono spedire il preventivo e la dichiarazione in lingua italiana all’indirizzo "),
                             ftext(RAMM.email), ftext(" e per conoscenza a "),
                             ftext(Richiedente..E.mail), ftext(" e jose.saporita@ipsp.cnr.it indicando nell’oggetto “Att.ne "), ftext(dott.ric), ftext(" "), ftext(Richiedente), ftext(": preventivo relativo all’avviso pubblico per "), ftext(la.fornitura), ftext(" di "), ftext(Prodotto), ftext("”.")), style = "Normal")
      }

      doc <- doc |>
        body_add_fpar(fpar(ftext("Individuazione dell'affidatario")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("L'individuazione dell'affidatario sarà operata discrezionalmente dalla Stazione Appaltante, nel caso in cui intenda procedere all’affidamento, a seguito dell'esame dei preventivi ricevuti entro la scadenza.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Non saranno presi in considerazione preventivi di importo superiore a quanto stimato dalla Stazione Appaltante.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("L’eventuale affidamento sarà concluso con l’operatore economico selezionato mediante affidamento diretto con trattativa diretta sul Mercato Elettronico della Pubblica Amministrazione (https://www.acquistinretepa.it/). A tal fine, l’operatore economico dovrà essere iscritto ed abilitato al bando “"),
                           ftext(beni), ftext("” del Mercato Elettronico, categorie “_____” oppure “_____”.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Obblighi dell’affidatario")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("L’operatore economico affidatario, con sede legale in Italia, sarà tenuto, prima dell’invio della lettera d’ordine, a fornire la seguente documentazione:")), style = "Normal")

      if(Importo.senza.IVA<40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Dichiarazione sostitutiva senza DGUE ai sensi del D.lgs. 36/2023;")), style = "Elenco punto")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("DGUE ai sensi del D.lgs. 36/2023;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("Comprovo assolvimento imposta di bollo;")), style = "Elenco punto")
      }

      doc <- doc |>
        body_add_fpar(fpar(ftext("Patto di integrità ai sensi del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Comunicazione conto corrente dedicato ai sensi dell’art. 3, comma 7 della Legge 136/2010 e s.m.i.;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Dichiarazione di cui al DPCM 187/1991.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("L'operatore economico straniero non residente in Italia, invece, sarà tenuto a fornire solo la seguente documentazione:")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Declaration on honour.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("La documentazione trasmessa dovrà essere sottoscritta digitalmente con firma qualificata da un legale rappresentante/procuratore in grado di impegnare l’operatore economico"), run_footnote(x=block_list(fpar(ftext(" Qualora l’operatore economico straniero fosse sprovvisto di firma digitale dovrà sottoscrivere la dichiarazione con firma autografa e allegare alla dichiarazione un documento d’identità in corso di validità.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript")), ftext(".")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Subappalto")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Non è consentito il subappalto delle prestazioni oggetto dell’affidamento, fermi restando i limiti e le condizioni di ricorso al subappalto per le prestazioni secondarie ed accessorie.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Chiarimenti")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Per eventuali ri chieste inerenti il servizio e chiarimenti di natura procedurale/amministrativa l’operatore economico dovrà rivolgersi "),
                           ftext(al.ric),ftext(" referente della Stazione Appaltante, "),
                           ftext(dott.ric),ftext(" "),ftext(Richiedente),
                           ftext(", all’indirizzo e-mail "),ftext(Richiedente..E.mail),ftext(".")),style = "Normal") |>
        body_add_fpar(fpar(ftext("Trattamento dei dati personali")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("I dati raccolti sono trattati e conservati ai sensi del Regolamento UE n. 2016/679 relativo alla protezione delle persone fisiche con riguardo al trattamento dei dati personali, nonché alla libera circolazione di tali dati, del decreto legislativo 30 giugno 2003, n. 196 recante il “Codice in materia di protezione dei dati personali” e s.m.i., del decreto della Presidenza del Consiglio dei ministri n. 148/21 e dei relativi atti di attuazione.")), style = "Normal") |>
        body_add_par("", style = "Normal") |>
        body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2") |>
        body_add_par("") |>
        body_end_section_continuous()

      b <- doc$officer_cursor$which +1
      e <- cursor_end(doc)
      e <- e$officer_cursor$which
      doc <- cursor_forward(doc)
      for(i in 1:(e-b)){
        doc <- body_remove(doc)
      }
      print(doc, target = paste0(pre.nome.file, "Avviso pubblico.docx"))

      ## Allegato ----
      doc <- doc.all |>
        cursor_begin() |>
        body_add_fpar(fpar(ftext("All’Istituto per la Protezione Sostenibile delle Piante")), style = "Destinatario", pos = "on") |>
        body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext("Oggetto:", fpt.b),
                           ftext(" indagine esplorativa di mercato volta a raccogliere preventivi finalizzati all’affidamento "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto),
                           ftext("” nell’ambito del progetto “"),
                           ftext(Progetto),
                           ftext("”"),
                           ftext(CUP1),
                           ftext(".")), style = "Oggetto maiuscoletto") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARAZIONE SOSTITUTIVA DELL’ATTO DI NOTORIETÀ")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("(resa ai sensi D.P.R. 28 dicembre 2000, n. 445)")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Il/La sottoscritto/a __________, nato/a a __________ il __________, codice fiscale __________, e residente a __________ in via __________, in qualità di legale rappresentante/procuratore della __________ con sede legale in via __________, CAP città (provincia), partita IVA __________, codice fiscale __________, telefono __________, PEC __________, e-mail __________, "),
                           ftext(" pienamente consapevole della responsabilità penale cui va incontro, ai sensi e per gli effetti dell’art. 76 D.P.R. 28 dicembre 2000, n. 445, in caso di dichiarazioni mendaci o di formazione, esibizione o uso di atti falsi ovvero di atti contenenti dati non più rispondenti a verità, ")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("di essere in possesso dei requisiti di cui all’avviso di indagine di mercato, e nello specifico:")), style = "Normal") |>
        body_add_fpar(fpar(ftext("requisiti di ordine generale di cui al Capo II, Titolo IV del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("requisiti d’idoneità professionale come specificato all’art. 100, comma 3 del D.lgs. n. 36/2023: iscrizione nel registro della camera di commercio, industria, artigianato e agricoltura o nel registro delle commissioni provinciali per l’artigianato o presso i competenti ordini professionali per un’attività pertinente anche se non coincidente con l’oggetto dell’appalto. All’operatore economico di altro Stato membro non residente in Italia è richiesto di dichiarare ai sensi del testo unico delle disposizioni legislative e regolamentari in materia di documentazione amministrativa, di cui al decreto del Presidente della Repubblica del 28 dicembre 2000, n. 445 di essere iscritto in uno dei registri professionali o commerciali di cui all’allegato II.11 del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("[nel caso di operatori economici residenti in Paesi terzi firmatari dell'AAP o di altri accordi internazionali di cui all'art. 69 del D.Lgs 36/2023]", fpt.i), ftext(" di essere iscritto in uno dei registri professionali e commerciali istituiti nel Paese in cui è residente;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Possesso di pregresse e documentate esperienze idonee all’esecuzione delle prestazioni contrattuali anche se non coincidenti con quelle oggetto dell’appalto.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Il sottoscritto dichiara, inoltre, di essere informato che, in conformità alla normativa vigente e in particolare al Regolamento GDPR 2016/679, i dati personali raccolti saranno trattati, anche con strumenti informatici, esclusivamente nell’ambito del procedimento per il quale la presente dichiarazione viene resa.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar("Firma digitale del legale rappresentante/procuratore", run_footnote(x=block_list(fpar(ftext("Per gli operatori economici italiani o stranieri residenti in Italia, la dichiarazione deve essere sottoscritta da un legale rappresentante ovvero da un procuratore del legale rappresentante, apponendo la firma digitale. Per gli operatori economici stranieri non residenti in Italia, la dichiarazione può essere sottoscritta dai medesimi soggetti apponendo la firma autografa ed allegando copia di un documento di identità del firmatario in corso di validità oppure con firma elettronica qualificata. Nel caso in cui la dichiarazione sia firmata da un procuratore del legale rappresentante, deve essere allegata copia conforme all’originale della procura oppure, nel solo caso in cui dalla visura camerale dell’operatore economico risulti l’indicazione espressa dei poteri rappresentativi conferiti con la procura, la dichiarazione sostitutiva resa dal procuratore/legale rappresentante sottoscrittore attestante la sussistenza dei poteri rappresentativi risultanti dalla visura.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_par("") |>
        body_end_section_continuous()

      b <- doc$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which
      for(i in 1:(e-b+1)){
        doc <- body_remove(doc)
      }
      doc <- headers_replace_all_text(doc, "*.*", "")
      print(doc, target = paste0(pre.nome.file, "Allegato.docx"))
      cat("

    Documenti '", pre.nome.file, "2 Avviso pubblico.docx' e '", pre.nome.file, "2.1 Allegato.docx' generati e salvati in ", pat)

      ## Dati mancanti ---
      manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, GAE, Richiedente..Luogo.di.nascita,
                             Richiedente..Codice.fiscale, Responsabile.progetto, Responsabile.progetto..Luogo.di.nascita, Responsabile.progetto..Codice.fiscale)
      manca <- as.data.frame(t(manca))
      colnames(manca) <- "val"
      manca$var <- rownames(manca)
      rownames(manca) <- NULL
      manca <- subset(manca, manca$val==trattini)
      len <- length(manca$val)
      if(len>0){
        manca <- manca$var
        manca <- paste0(manca, ",")
        manca[len] <- sub(",$", "\\.", manca[len])
        cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
        cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
      }
    }
  }

  # Genera DaC ----
  dac <- function(){

    if(Fornitore==fornitore.uscente){
      cat(paste0(
        "***** ATTENZIONE *****\n",
        Fornitore, " è il fornitore uscente.\n",
        "L'ultimo ordine (n° ", ordine.uscente, ") per questa categoria merceologica (prime tre cifre del CPV: ", cpv.usente, ") è stato affidato a questo operatore economico per l'acquisto di '", prodotto.uscente, "' e un importo di € ", importo.uscente, ".\n"))
      if(Rotazione.fornitore=="Non è il contraente uscente"){
        cat("In FluOr è stato erroneamente indicato 'Non è il contraente uscente'. Si prega di apportare la dovuta correzione.\n")
      }else if(Rotazione.fornitore=="Particolare struttura del mercato"){
        cat("L'ordine può procedere poichè è stato indicato 'Particolare struttura del mercato'.\n")
      }else if(Rotazione.fornitore=="Importo <5.000€" & Importo.senza.IVA.num<5000){
        cat("L'ordine può procedere poichè è stata specificata la deroga alla rotazione dei fornitori per ordini <5.000 €.\n")
      }else if(Rotazione.fornitore=="Importo <5.000€" & Importo.senza.IVA.num>=5000){
        cat("E' stata specificata la deroga alla rotazione dei fornitori per ordini <5.000 €, ma l'ordine è superiore a questo importo. Si prega di apportare la dovuta correzione.\n")
      }
      cat("*********************\n",
          " Premere INVIO per proseguire")
      readline()
    }

    doc <- doc.dac |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE) |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_par("PROVVEDIMENTO", style = "heading 1", pos = "on") |>
      body_add_fpar(fpar(ftext("OGGETTO:", fpt.b),
                         ftext(" Decisione di contrattare per l’affidamento diretto, ai sensi dell’art. 50, comma 1, lett. b) del D.lgs. n. 36/2023, "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” nell’ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", ordine CNR-IPSP-"),
                         ftext(sede),
                         ftext(" N° "),
                         ftext(ordine),
                         ftext(y),
                         ftext(".")), style = "Oggetto") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 31 dicembre 2009 n. 213, recante “Riordino del Consiglio Nazionale delle Ricerche in attuazione dell’articolo 1 della Legge 27 settembre 2007, n. 165“;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 25 novembre 2016 n. 218, recante “Semplificazione delle attività degli enti pubblici di ricerca ai sensi dell'articolo 13 della legge 7 agosto 2015, n. 124”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 7 agosto 1990, n. 241 recante “Nuove norme in materia di procedimento amministrativo e di diritto di accesso ai documenti amministrativi” pubblicata sulla Gazzetta Ufficiale n. 192 del 18/08/1990 e s.m.i.;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Regolamento di Organizzazione e Funzionamento del Consiglio Nazionale delle Ricerche emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 01/08/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il D.lgs. 31 marzo 2023, n. 36 rubricato “Codice dei Contratti Pubblici in attuazione dell’articolo 1 della legge 21 giugno 2022, n. 78, recante delega al Governo in materia di contratti pubblici”, pubblicato sul Supplemento Ordinario n. 12 della GU n. 77 del 31 marzo 2023 (nel seguito per brevità “Codice”);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 59 del Regolamento di Amministrazione, Contabilità e Finanza del Consiglio Nazionale delle Ricerche rubricato “Decisione di contrattare” – DPCNR del 04 maggio 2005 prot. 0025034 pubblicato sulla G.U.R.I. n. 124 del 30/05/2005 – Supplemento Ordinario n. 101;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 10.000,000 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                         ftext(nomina.RAMM)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
                         ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 59 del Regolamento di amministrazione, contabilità e finanza del CNR (Decreto del Presidente del CNR prot. 25034 del 4/5/2005);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 6 novembre 2012, n. 190 recante “Disposizioni per la prevenzione e la repressione della corruzione e dell’illegalità nella pubblica amministrazione” pubblicata sulla G.U.R.I. n. 265 del 13/11/2012;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 14 marzo 2013, n. 33 recante “Riordino della disciplina riguardante il diritto di accesso civico e gli obblighi di pubblicità, trasparenza e diffusione di informazioni da parte delle pubbliche amministrazioni” pubblicato sulla Gazzetta Ufficiale n. 80 del 05/04/2013 e successive modifiche introdotte dal d.lgs. 25 maggio 2016 n. 97;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Codice di comportamento dei dipendenti del Consiglio Nazionale delle Ricerche approvato con delibera del Consiglio di Amministrazione n° 137/2017;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il vigente Piano triennale per la prevenzione della corruzione e della trasparenza (PTPCT), adottato con delibera del Consiglio di Amministrazione del Consiglio Nazionale delle Ricerche ai sensi della legge 6 novembre 2012 n. 190;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 23 dicembre 1999 n 488 e s.m.i., recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2000)”, ed in particolare l'articolo 26;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 27 dicembre 2006 n. 296, recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2007)”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 24 dicembre 2007 n. 244 e s.m.i., recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2008)”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il decreto-legge 7 maggio 2012 n. 52, convertito dalla legge 6 luglio 2012 n. 94 recante “Disposizioni urgenti per la razionalizzazione della spesa pubblica”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il decreto-legge 6 luglio 2012 n. 95, convertito con modificazioni dalla legge 7 agosto 2012 n. 135, recante “Disposizioni urgenti per la revisione della spesa pubblica con invarianza dei servizi ai cittadini”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la richiesta di acquisto prot. "),
                         ftext(Prot..RAS), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” (pagina web dedicata al ciclo di vita del contratto pubblico "),
                         ftext(Pagina.web),
                         ftext("), nell’ambito delle attività previste dal progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", mediante affidamento diretto all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(") per un importo stimato di "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA, il cui preventivo è "),
                         ftext(preventivo.individuato)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 50, comma 1, lettera b) del Codice, il quale consente, per affidamenti di contratti di servizi e forniture, ivi compresi i servizi di ingegneria e architettura e l'attività di progettazione di importo inferiore ad euro 140.000,00, di procedere ad affidamento diretto, anche senza consultazione di più operatori economici;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(", altresì, che la scelta dell’Operatore Economico deve essere effettuata assicurando che i soggetti individuati siano in possesso di documentate esperienze pregresse idonee all’esecuzione delle prestazioni contrattuali, anche individuati tra gli iscritti in elenchi o albi istituiti dalla stazione appaltante;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Comunicato del Presidente ANAC del 10 gennaio 2024 con cui sono state diramate indicazioni di carattere transitorio sull’applicazione delle disposizioni del codice dei contratti pubblici in materia di digitalizzazione degli affidamenti di importo inferiore a euro 5.000,00"),
                         ftext(" [valido fino al 30.09.2024];", fpt.i)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VALUTATA", fpt.b), ftext(" l’opportunità, in ottemperanza alla suddetta normativa, di procedere ad affidamento diretto all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(") mediante provvedimento contenente gli elementi essenziali descritti nell’art. 17, comma 2, del Codice, tenuto conto che il medesimo è in possesso di documentate esperienze pregresse idonee all’esecuzione della prestazione contrattuale;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO ", fpt.b),
                         ftext(rotazione.individuata)), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che dal 1° gennaio 2024 è fatto obbligo di utilizzare piattaforme di approvvigionamento digitale certificate (e-procurement) per svolgere le procedure di affidamento e di esecuzione dei contratti pubblici, a norma degli artt. 25 e 26 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che la stazione appaltante ai sensi dell’art. 48 comma 2 del Codice, ha accertato che il presente appalto non presenta un interesse transfrontaliero certo per cui non segue le procedure ordinarie di cui alla parte IV del Libro II;")), style = "Normal")

    if(Motivo.fuori.MePA!="No"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(ICT.testo)), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che ai sensi dell’art. 53 comma 1 del Codice non sussistono particolari ragioni per la richiesta di garanzia provvisoria;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il bilancio di previsione del Consiglio Nazionale delle Ricerche per l'esercizio finanziario 2024, approvato dal Consiglio di Amministrazione con deliberazione n° 371/2023 del 28/11/2023, Verb. 488;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b), ftext(" la disponibilità finanziaria per la copertura della spesa sui fondi del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", allocati sul GAE "),
                         ftext(GAE),
                         ftext(", voce del piano "),
                         ftext(Voce.di.spesa),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che vi sono i presupposti normativi e di fatto per acquisire "), ftext(la.fornitura), ftext(" in oggetto, nel rispetto dei principi generali enunciati nel Codice;")), style = "Normal") |>
      body_add_par("DISPONE", style = "heading 2") |>
      body_add_fpar(fpar(ftext("DI PROCEDERE", fpt.b), ftext(" all’acquisizione "), ftext(della.fornitura), ftext(", mediante affidamento diretto ai sensi dell’art. 50, comma 1, lett. b) del Codice, all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" con sede legale in "),
                         ftext(Fornitore..Sede),
                         ftext(", C.F. e P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", per un importo complessivo pari a euro "),
                         ftext(Importo.senza.IVA),
                         ftext(", al netto dell’IVA e di altre imposte e contributi di legge, ritenuto congruo in relazione alle condizioni di mercato;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("DI NOMINARE ", fpt.b),
                         ftext(il.dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(" Responsabile Unico del Progetto il quale, ai sensi dell’art. 15 del Codice, dovrà:")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("svolgere tutte le attività indicate nell’allegato I.2 del Codice, o che siano comunque necessarie ove non di competenza di altri organi;")), style = "Elenco liv2") |>
      body_add_fpar(fpar(ftext("vigilare sullo svolgimento delle fasi di affidamento ed esecuzione "), ftext(della.fornitura), ftext(" in parola, provvedendo a creare le condizioni affinché il processo di acquisto risulti condotto in modo unitario rispetto alle esigenze ed ai costi indicati nel presente atto, in conformità a qualsiasi altra disposizione di legge e di regolamento in materia ivi incluso l’accertamento dei requisiti di carattere generale e tecnico-professionali, ove richiesti, in capo all’operatore economico individuato; ")), style = "Elenco liv2")

    if(Tipo.ordine=="Trattativa diretta MePA" | Tipo.ordine=="RDO MePA" | Tipo.ordine=="Ordine diretto MePA"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("procedere alla prenotazione del Codice Identificativo Gara (CIG) tramite la piattaforma di approvvigionamento digitale certificata e a tutti gli altri adempimenti previsti dalla normativa vigente;")), style = "Elenco liv2")
    }else{
      doc <- doc |>
      body_add_fpar(fpar(ftext("procedere alla prenotazione del Codice Identificativo Gara (CIG) tramite la piattaforma contratti pubblici (PCP) dell’ANAC e a tutti gli altri adempimenti previsti dalla normativa vigente;")), style = "Elenco liv2")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("rilasciare apposita dichiarazione, rispetto al ruolo ricoperto ed alle funzioni svolte, nella quale attesti di non trovarsi in alcuna delle situazioni di conflitto di interessi, anche potenziale, di cui all’art. 16 del D.lgs. n. 36/2023;")), style = "Elenco liv2") |>
      #body_add_fpar(fpar(ftext("DI INDIVIDUARE", fpt.b), ftext(" ai sensi dell’art. 15, comma 6 del Codice, il dott. Nicola Centorame in qualità di supporto al RUP;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" che l'affidamento di cui al presente provvedimento sia soggetto all’applicazione delle norme contenute nella legge n. 136/2010 e s.m.i. e che il pagamento venga disposto entro 30 giorni dall’emissione certificato di regolare esecuzione;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" che, ai sensi dell'art. 53 del Codice l'affidatario sia esonerato dalla costituzione della garanzia definitiva in quanto l'ammontare garantito sarebbe di importo così esiguo da non costituire reale garanzia per la stazione appaltante, determinando esclusivamente un appesantimento del procedimento;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("DI ASSUMERE", fpt.b), ftext(" l’impegno provvisorio di spesa n. "),
                         ftext(N..impegno.di.spesa),
                         ftext(" per un importo pari a euro "),
                         ftext(Importo.con.IVA),
                         ftext(" comprensivo di IVA sui fondi del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", allocati sul GAE "),
                         ftext(GAE),
                         ftext(", voce del piano "),
                         ftext(Voce.di.spesa),
                         ftext(", a favore del codice terzo registrato in SIGLA con il n. "),
                         ftext(Fornitore..Codice.terzo.SIGLA),
                         ftext(";")), style = "Elenco liv1")

    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" che l'avvio dell'esecuzione del contratto o la sottoscrizione dello stesso/l’invio della lettera d’ordine siano subordinati all'esito della verifica dei requisiti di ordine generale, e speciale se previsti, senza rilevare cause ostative;")), style = "Elenco liv1") |>
        body_add_fpar(fpar(ftext("DI IMPEGNARE", fpt.b), ftext(" la spesa per un importo pari a € 35,00 sui fondi del già citato progetto, allocati sul GAE P___, voce del piano 13096 “Pubblicazione bandi di gara” per la contribuzione ANAC;")), style = "Elenco liv1")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI SOTTOPORRE", fpt.b), ftext(" la lettera d’ordine alla condizione risolutiva in caso di accertamento della carenza dei requisiti di ordine generale;")), style = "Elenco liv1")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("DI PROCEDERE", fpt.b), ftext(" alla pubblicazione del presente provvedimento ai sensi del combinato disposto dell’Art. 37 del d.lgs. 14 marzo 2013, n. 33 e dell’art. 20 del Codice;")), style = "Elenco liv1") |>
      body_add_par("DICHIARA", style = "heading 2") |>
      body_add_par("l’insussistenza a proprio carico di situazioni di conflitto di interesse di cui all’art. 16 del Codice.", style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_par("Visto di regolarità contabile", style = "Firma 1") |>
      body_add_par("Il Responsabile Amministrativo", style = "Firma 1") |>
      body_add_fpar(fpar(ftext("("), ftext(RAMM), ftext(")")), style = "Firma 1") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    print(doc, target = paste0(pre.nome.file, "4 Decisione a contrattare.docx"))
    print(doc, target = paste0(pre.nome.file, "4 Decisione a contrattare per URP.docx"))

    #cat("\014")
    #cat(rep("\n", 20))
    cat("

    Documento '", pre.nome.file, "4 Decisione a contrattare.docx' e 'Decisione a contrattare per URP.docx' generati e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, GAE, Richiedente..Luogo.di.nascita,
                           Richiedente..Codice.fiscale, Responsabile.progetto, Responsabile.progetto..Luogo.di.nascita, Responsabile.progetto..Codice.fiscale)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
        **********************")
    }

    ## DURC
    if(Fornitore..DURC.scadenza!=trattini){
      if(Fornitore..DURC.scadenza<=da){
        cat("
    ***** ATTENZIONE *****
    Il DURC di", Fornitore, "è scaduto il giorno", Fornitore..DURC.scadenza, "
    **********************")
      }
    }
  }

  # Genera Provv. impegno ----
  provv_imp <- function(){
    doc <- doc.prov.imp |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE) |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("CdR 121.001.000 IPSP ", fpt.b), ftext(sede2, fpt.b)), style = "Normal") |>
      body_add_fpar(fpar(ftext("PROVVEDIMENTO DI IMPEGNO DELLA")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("LETTERA D'ORDINE CNR-IPSP-"), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1") |>
      body_add_fpar(fpar(firma.RSS), style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Organizzazione e Funzionamento del CNR emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 01/08/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il DP CNR n. 0025034 in data 4 maggio 2005 concernente il Regolamento di amministrazione, contabilità e finanza del Consiglio Nazionale delle Ricerche e in particolare l’art. 28 “Impegno”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento CNR n. 114 del 30/10/2013 (prot. n. 0065484) relativo alla costituzione dell’Istituto per la Protezione Sostenibile delle Piante con successivi provvedimenti del Presidente n. 120 del 07/10/2014 (prot. n. 72102) e n. 2 del 11/01/2019 di conferma e sostituzione del precedente atto costitutivo;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 10.000,000 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                         ftext(nomina.RAMM)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
                         ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 59 del Regolamento di amministrazione, contabilità e finanza del CNR (Decreto del Presidente del CNR prot. 25034 del 4/5/2005);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la richiesta di acquisto prot. "),
                         ftext(Prot..RAS), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” (pagina web dedicata al ciclo di vita del contratto pubblico "),
                         ftext(Pagina.web),
                         ftext("), nell’ambito delle attività previste dal progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", mediante affidamento diretto all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", codice terzo SIGLA "),
                         ftext(Fornitore..Codice.terzo.SIGLA),
                         ftext(") per un importo stimato di "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la verifica effettuata dal Responsabile Amministrativo della copertura finanziaria (art. 28, comma 2 Regolamento di contabilità);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b),
                         ftext(" che la fornitura in oggetto è funzionalmente destinata all’attività di ricerca;")), style = "Normal") |>
      body_add_par("DISPONE", style = "heading 2") |>
      body_add_fpar(fpar(ftext("l’assunzione dell'impegno di spesa n° "),
                         ftext(N..impegno.di.spesa),
                         ftext(" di "),
                         ftext(Importo.con.IVA),
                         ftext(" IVA inclusa, con imputazione sulla voce di spesa "),
                         ftext(Voce.di.spesa),
                         ftext(", GAE "),
                         ftext(GAE),
                         ftext(", progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(";")), style = "Elenco punto")

      # if(CUI!=trattini){
      # doc <- doc |>
      #   body_add_fpar(fpar(ftext("CUI: "), ftext(CUI), ftext(";")), style = "Elenco punto liv2")
      # }

      if(Tipo.acquisizione=='Beni'){
      if(Inventariabile=='Inventariabile'){
        doc <- doc |>
          body_add_fpar(fpar(ftext("di inventariare il bene e nominare "),
                             ftext(assegna),
                             ftext(" "),
                             ftext(il.dott.ric), ftext(" "), ftext(Richiedente), ftext(".")), style = "Elenco punto")
      }
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("di non inventariare il bene in quanto trattasi di materiale di consumo.")), style = "Elenco punto")
    }

    doc <- doc |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    print(doc, target = paste0(pre.nome.file, "3 Provv. impegno.docx"))
    cat("

    Documento '", pre.nome.file, "3 Provv. impegno.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Fornitore, Fornitore..P.IVA, Fornitore..Codice.terzo.SIGLA, N..impegno.di.spesa, Importo.con.IVA, Voce.di.spesa, GAE, Richiedente)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
    }
  }

  # Genera Richiesta pagina web ----
  pag <- function(){
    doc <- doc.pag |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("All'"), ftext("Unità Relazioni con il Pubblico", fpt.b)), style = "Destinatario") |>
      body_add_fpar(fpar(ftext("del CNR")), style = "Destinatario 2") |>
      body_add_par("", style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("OGGETTO", fpt.bi), ftext(": procedura per l’affidamento diretto, ai sensi dell’art. 50, comma 1, lettera b del D.lgs. n. 36/2023, ", fpt.i),
                         ftext(della.fornitura, fpt.i), ftext(" di “", fpt.i),
                         ftext(Prodotto, fpt.i),
                         ftext("” - "),
                         ftext("Richiesta pagina dedicata al ciclo di vita del contratto pubblico.", fpt.bi)), style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("Con riferimento alla procedura in oggetto ed ai fini della sua pubblicazione si chiede la cortesia di procedere all’apertura della pagina dedicata al ciclo di vita dell’affidamento di cui sopra.")), style = "Normal") |>
      body_add_fpar(fpar(ftext("Si ringrazia per la cortese collaborazione.")), style = "Normal") |>
      body_add_fpar(fpar(ftext("Cordiali saluti.")), style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("Il Responsabile Unico del Progetto")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(Dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    print(doc, target = paste0(pre.nome.file, "1.1 Richiesta pagina web.docx"))
    cat("

    Documento '", pre.nome.file, "1.1 Richiesta pagina web.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, RUP)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
    }
  }

  # Genera DocOE ----
  docoe <- function(){
    inpt.oe <- 1
    if(ultimi.recente>0 & ultimi.recente<180){
      cat(paste0("

      I documenti dell'operatore economico ", Fornitore, " sono già stati richiesti meno di 6 mesi fa (prot. ", ultimi.prot, ") in occasione dell'ordine n° ", ultimi.ordine, y,".
Si vuole generare ugualmente i documenti dell'operatore economico per richiederli nuovamente?
  1: Sì
  2: No"))
      inpt.oe <- readline()
    }

    if(inpt.oe==1){
      if(Fornitore..Nazione=="Italiana"){

        ## Patto d'integrità ----
        doc <- doc.pi |>
          cursor_begin() |>
          headers_replace_all_text("CAMPO.CIG.PATTO", paste0("CIG ", CIG, " (", Pagina.web, ")"), only_at_cursor = FALSE) |>
          cursor_reach("CAMPO.DELLA.FORNITURA.PATTO") |>
          body_replace_all_text("CAMPO.DELLA.FORNITURA.PATTO", paste0(della.fornitura, " di “", Prodotto, "” (CIG ", CIG, ", ", Pagina.web, "), nell'ambito del progetto ", Progetto1), only_at_cursor = TRUE) |>
          cursor_reach("CAMPO.FORNITORE.PATTO") |>
          body_replace_all_text("CAMPO.FORNITORE.PATTO", paste0("L'operatore economico ", Fornitore, " (di seguito Operatore Economico) con sede legale in ", Fornitore..Sede, ", C.F./P.IVA ", as.character(Fornitore..P.IVA), ", rappresentato da ", Fornitore..Rappresentante.legale, " in qualità di ", tolower(Fornitore..Ruolo.rappresentante), ","), only_at_cursor = TRUE) |>
          cursor_reach("CAMPO.FIRMA.RSS.PATTO") |>
          body_replace_all_text("CAMPO.FIRMA.RSS.PATTO", firma.RSS, only_at_cursor = TRUE) |>
          headers_replace_all_text("CAMPO.CIG.PATTO", paste0("CIG ", CIG, " (", Pagina.web, ")"), only_at_cursor = FALSE)

        b <- cursor_begin(doc)
        b <- b$officer_cursor$which
        e <- cursor_reach(doc, "PATTO DI INTEGRIT")
        e <- e$officer_cursor$which -1
        doc <- cursor_begin(doc)
        doc <- cursor_forward(doc)
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }
        doc <- cursor_reach(doc, "PATTO DI INTEGRIT")
        doc <- cursor_backward(doc)
        doc <- body_remove(doc)

        b <- cursor_reach(doc, "Oggetto: Comunicazione c/c")
        b <- b$officer_cursor$which
        e <- cursor_end(doc)
        e <- e$officer_cursor$which +5
        doc <- cursor_reach(doc, "Oggetto: Comunicazione c/c")
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }
        doc <- cursor_backward(doc)
        doc <- body_remove(doc)
        print(doc, target = paste0(pre.nome.file, "5.1 Patto di integrità.docx"))

        ## CC dedicato ----
        doc <- doc.cc
        b <- cursor_begin(doc)
        b <- b$officer_cursor$which
        e <- cursor_reach(doc, "Oggetto: Comunicazione c/c dedicato ")
        e <- e$officer_cursor$which -3
        doc <- cursor_begin(doc)
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }

        b <- cursor_reach(doc, "DICHIARAZIONE POSSESSO REQUISITI DI PARTECIPAZIONE E DI QUALIFICAZIONE")
        b <- b$officer_cursor$which
        e <- cursor_end(doc)
        e <- e$officer_cursor$which
        doc <- cursor_reach(doc, "DICHIARAZIONE POSSESSO REQUISITI DI PARTECIPAZIONE E DI QUALIFICAZIONE")
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }
        doc <- cursor_reach(doc, "del legale rappresentante/procuratore")
        doc <- cursor_forward(doc)
        doc <- body_remove(doc)
        print(doc, target = paste0(pre.nome.file, "5.2 Comunicazione cc dedicato.docx"))

        ## DPCM ----
        doc <- doc.dpcm
        b <- cursor_begin(doc)
        b <- b$officer_cursor$which
        e <- cursor_reach(doc, "CAMPO.INT.DOC.DPCM")
        e <- e$officer_cursor$which
        doc <- cursor_begin(doc)
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }

        doc <- doc |>
          cursor_reach("CAMPO.INT.DOC.DPCM") |>
          body_replace_all_text("CAMPO.INT.DOC.DPCM", int.doc, only_at_cursor = TRUE)

        b <- cursor_reach(doc, "DECLARATION ON HONOUR")
        b <- b$officer_cursor$which
        e <- cursor_end(doc)
        e <- e$officer_cursor$which
        doc <- cursor_reach(doc, "DECLARATION ON HONOUR")
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }
        doc <- cursor_reach(doc, "del legale rappresentante/procuratore")
        doc <- cursor_forward(doc)
        doc <- body_remove(doc)
        print(doc, target = paste0(pre.nome.file, "5.3 Dichiarazione DPCM 187.docx"))
        #cat("\014")
        #cat(rep("\n", 20))
        #cat("\014")
        cat("

    Documenti '", pre.nome.file, "5.1 Patto di integrità.docx', '5.2 Comunicazione cc dedicato.docx' e '5.3 Dichiarazione DPCM 187.docx' generati e salvati in ", pat)

        ## Dati mancanti ---
        manca <- dplyr::select(sc, Fornitore, Fornitore..Sede, Fornitore..P.IVA, Prodotto, Progetto, Pagina.web)
        manca <- as.data.frame(t(manca))
        colnames(manca) <- "val"
        manca$var <- rownames(manca)
        rownames(manca) <- NULL
        manca <- subset(manca, manca$val==trattini)
        len <- length(manca$val)
        if(len>0){
          manca <- manca$var
          manca <- paste0(manca, ",")
          manca[len] <- sub(",$", "\\.", manca[len])
          cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
          cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
        }
          if(Importo.senza.IVA.num<40000){
            ## Part.Qual. ----
            doc <- doc.part.qual
            b <- cursor_begin(doc)
            b <- b$officer_cursor$which
            e <- cursor_reach(doc, "DICHIARAZIONE POSSESSO REQUISITI DI PARTECIPAZIONE E DI QUALIFICAZIONE")
            e <- e$officer_cursor$which
            doc <- cursor_begin(doc)
            for(i in 1:(e-b)){
              doc <- body_remove(doc)
            }

            doc <- doc |>
              cursor_begin() |>
              cursor_reach("CAMPO.INT.DOC") |>
              body_replace_all_text("CAMPO.INT.DOC", int.doc, only_at_cursor = TRUE) |>
              headers_replace_all_text("“Dichiarazione di cui al DPCM 187/1991”", "")

            b <- cursor_reach(doc, "DICHIARAZIONE RELATIVA AL POSSESSO DEI REQUISITI DI QUALIFICAZIONE")
            b <- b$officer_cursor$which
            e <- cursor_end(doc)
            e <- e$officer_cursor$which
            doc <- cursor_reach(doc, "DICHIARAZIONE RELATIVA AL POSSESSO DEI REQUISITI DI QUALIFICAZIONE")
            for(i in 1:(e-b)){
              doc <- body_remove(doc)
            }
            doc <- cursor_backward(doc)
            doc <- cursor_backward(doc)
            doc <- body_remove(doc)
            print(doc, target = paste0(pre.nome.file, "5.4 Dichiarazione possesso requisiti di partecipazione e qualificazione.docx"))

            cat("
    Documento '", pre.nome.file, "5.4 Dichiarazione possesso requisiti di partecipazione e qualificazione.docx' generato e salvato in ", pat)
          }

            if(Importo.senza.IVA.num>=40000){
              ## Qual. ----
              doc <- doc.qual
              b <- cursor_begin(doc)
              b <- b$officer_cursor$which
              e <- cursor_reach(doc, "DICHIARAZIONE RELATIVA AL POSSESSO DEI REQUISITI DI QUALIFICAZIONE")
              e <- e$officer_cursor$which
              doc <- cursor_begin(doc)
              for(i in 1:(e-b)){
                doc <- body_remove(doc)
              }

              doc <- doc |>
                cursor_begin() |>
                cursor_reach("CAMPO.INT.DOC") |>
                body_replace_all_text("CAMPO.INT.DOC", int.doc, only_at_cursor = TRUE) |>
                headers_replace_all_text("“Dichiarazione di cui al DPCM 187/1991”", "")

              b <- cursor_reach(doc, "CAMPO.INT.DOC.DPCM")
              b <- b$officer_cursor$which
              e <- cursor_end(doc)
              e <- e$officer_cursor$which
              doc <- cursor_reach(doc, "CAMPO.INT.DOC.DPCM")
              for(i in 1:(e-b)){
                doc <- body_remove(doc)
              }
              b <- cursor_reach(doc, "del legale rappresentante/procuratore")
              b <- b$officer_cursor$which
              e <- cursor_end(doc)
              e <- e$officer_cursor$which -1
              doc <- cursor_reach(doc, "del legale rappresentante/procuratore")
              doc <- cursor_forward(doc)
              for(i in 1:(e-b)){
                doc <- body_remove(doc)
              }
              print(doc, target = paste0(pre.nome.file, "5.4 Dichiarazione possesso requisiti di qualificazione.docx"))

              cat("
    Documento '", pre.nome.file, "5.4 Dichiarazione possesso requisiti di qualificazione.docx' generato e salvato in ", pat)

              ## AUS ----
              doc <- doc.aus
              b <- cursor_begin(doc)
              b <- b$officer_cursor$which
              e <- cursor_reach(doc, "DICHIARAZIONE SOSTITUTIVA DEL SOGGETTO AUSILIARIO")
              e <- e$officer_cursor$which
              doc <- cursor_begin(doc)
              for(i in 1:(e-b)){
                doc <- body_remove(doc)
              }

              doc <- doc |>
                cursor_reach("CAMPO.INT.DOC") |>
                body_replace_all_text("CAMPO.INT.DOC", int.doc, only_at_cursor = TRUE)

              print(doc, target = paste0(pre.nome.file, "5.5 Dichiarazione del soggetto ausiliario.docx"))

              cat("
    Documento '", pre.nome.file, "5.5 Dichiarazione del soggetto ausiliario.docx' generato e salvato in ", pat)
            }
      }else{
        ## Declaration on honour ----
        doc <- doc.doh
        b <- cursor_begin(doc)
        b <- b$officer_cursor$which
        e <- cursor_reach(doc, "DECLARATION ON HONOUR")
        e <- e$officer_cursor$which
        doc <- cursor_begin(doc)
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }

        b <- cursor_reach(doc, "DICHIARAZIONE SOSTITUTIVA DEL SOGGETTO AUSILIARIO")
        b <- b$officer_cursor$which
        e <- cursor_end(doc)
        e <- e$officer_cursor$which
        doc <- cursor_reach(doc, "DICHIARAZIONE SOSTITUTIVA DEL SOGGETTO AUSILIARIO")
        for(i in 1:(e-b)){
          doc <- body_remove(doc)
        }
        doc <- cursor_reach(doc, "Signature")
        doc <- cursor_forward(doc)
        doc <- body_remove(doc)

        print(doc, target = paste0(pre.nome.file, "5.7 Declaration on honour.docx"))
        cat("\014")
        #cat(rep("\n", 20))
        cat("\014")
        cat("

    Documento '", pre.nome.file, "5.7 Declaration on honour.docx' generato e salvato in ", pat)
      }

        if(Importo.senza.IVA.num>=40000){
          ## Bollo ----
          doc <- doc.bollo
          b <- cursor_begin(doc)
          b <- b$officer_cursor$which
          e <- cursor_reach(doc, "DICHIARAZIONE POSSESSO REQUISITI DI PARTECIPAZIONE E DI QUALIFICAZIONE")
          e <- e$officer_cursor$which
          doc <- cursor_begin(doc)
          for(i in 1:(e-b)){
            doc <- body_remove(doc)
          }

          b <- cursor_reach(doc, "DICHIARAZIONE POSSESSO REQUISITI DI PARTECIPAZIONE E DI QUALIFICAZIONE")
          b <- b$officer_cursor$which
          e <- cursor_end(doc)
          e <- e$officer_cursor$which
          doc <- cursor_reach(doc, "DICHIARAZIONE POSSESSO REQUISITI DI PARTECIPAZIONE E DI QUALIFICAZIONE")
          for(i in 1:(e-b)){
            doc <- body_remove(doc)
          }
          doc <- doc |>
            headers_replace_all_text(".*", "", only_at_cursor = TRUE) |>
            cursor_begin() |>
            body_add_fpar(fpar(ftext("")), style = "Normal") |>
            body_add_fpar(fpar(ftext("COMPROVA IMPOSTA DI BOLLO", fpt.b), run_footnote(x=block_list(fpar(ftext(" Sono esenti i contratti di importo inferiore a 40.000,00 euro. Pagamento di 40 euro, per i contratti di importo maggiore o uguale a 40 mila e inferiore a 150mila euro.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Normal") |>
            body_add_fpar(fpar(ftext("Affidamento diretto, ai sensi dell’art. 50 del D.Lgs. N. 36/2023, "),
                               ftext(della.fornitura), ftext(" di “"),
                               ftext(Prodotto),
                               ftext("” (CIG "),
                               ftext(CIG),
                               ftext(CUI1),
                               ftext(", "), ftext(Pagina.web),
                               ftext("), nell'ambito del progetto “"),
                               ftext(Progetto),
                               ftext("”"),
                               ftext(CUP1),
                               ftext(ordine.trattativa.scelta),
                               ftext(", ordine CNR-IPSP-"),
                               ftext(sede),
                               ftext(" N° "),
                               ftext(ordine),
                               ftext(y),
                               ftext(".")), style = "Oggetto maiuscoletto") |>
            body_add_fpar(fpar(ftext("")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Il sottoscritto ____________________")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Codice fiscale ____________________")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Nella sua qualità di:")), style = "Normal") |>
            body_add_fpar(fpar(ftext("   Titolare o Legale rappresentante")), style = "Normal") |>
            body_add_fpar(fpar(ftext("   Procuratore")), style = "Normal") |>
            body_add_fpar(fpar(ftext("del concorrente ____________________")), style = "Normal") |>
            body_add_fpar(fpar(ftext("")), style = "Normal") |>
            body_add_fpar(fpar(ftext("consapevole che le false dichiarazioni, la falsità degli atti e l’uso di atti falsi sono puniti ai sensi del codice penale (Artt. 75 e 76 del D.P.R. 445/2000)")), style = "Normal") |>
            body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2")

          if(Fornitore..Nazione=="Italiana"){
            doc <- doc |>
              body_add_fpar(fpar(ftext("che l’imposta di bollo è stata assolta in modalità telematica, utilizzando il modello “F24 Versamenti con elementi identificativi” (F24 ELIDE) e che la relativa quietanza è allegata al documento a comprova del versamento;")), style = "Elenco punto")
          }else{
            doc <- doc |>
              body_add_fpar(fpar(ftext("che l’imposta di bollo è stata assolta mediante bonifico bancario pertanto si allega copia dello stesso a comprova del versamento;")), style = "Elenco punto")
          }

          doc <- doc |>
            body_add_fpar(fpar(ftext("di essere a conoscenza che la Stazione appaltante potrà effettuare controlli sui documenti presentati e pertanto si impegna a conservare il presente documento fino al termine di decadenza triennale previsto per l’accertamento da parte dell’Amministrazione finanziaria (Art. 37 D.P.R. N° 642/1972) e a renderlo disponibile ai fini dei successivi controlli.")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Firma digitale"), run_footnote(x=block_list(fpar(ftext("   Per gli operatori economici italiani o stranieri residenti in Italia, la dichiarazione deve essere sottoscritta da un legale rappresentante ovvero da un procuratore3 del legale rappresentante, apponendo la firma digitale. Per gli operatori economici stranieri non residenti in Italia, la dichiarazione può essere sottoscritta dai medesimi soggetti apponendo la firma autografa ed allegando copia di un documento di identità del firmatario in corso di validità.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript")),
                               ftext(" del legale rappresentante/procuratore"), run_footnote(x=block_list(fpar(ftext(" Nel caso in cui la dichiarazione sia firmata da un procuratore del legale rappresentante deve essere allegata copia conforme all’originale della procura oppure nel solo caso in cui dalla visura camerale dell’operatore economico risulti l’indicazione espressa dei poteri rappresentativi conferiti con la procura, la dichiarazione sostitutiva resa dal procuratore/legale rappresentante sottoscrittore attestante la sussistenza dei poteri rappresentativi risultanti dalla visura.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
            body_add_break()|>
            body_add_fpar(fpar(ftext("Indicazioni per il pagamento dell’imposta di bollo sul contratto", fpt.b)), style = "Normal") |>
            body_add_fpar(fpar(ftext("")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Ai sensi dell’art. 18, comma 10 del Codice, con la tabella di cui all’allegato I.4 al Codice è individuato il valore dell’imposta di bollo che l’aggiudicatario assolve una tantum al momento della stipula del contratto e in proporzione al valore dello stesso.")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Il pagamento dell’imposta di bollo sul contratto deve essere effettuato con versamento tramite modello “F24 Versamenti con elementi identificativi” (F24 ELIDE) ai sensi del provvedimento del Direttore dell’Agenzia delle Entrate (AdE) Prot. n. 240013/2023 del 28/06/2023. Le modalità di compilazione del modello F24 ELIDE sono individuate nella Risoluzione AdE n. 37/E del 28/06/2023.")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Nel caso specifico l’imposta da versare con le modalità di cui sopra è pari a € 40,00; segue il riepilogo dei campi da compilare e dei dati da utilizzare:")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Sezione contribuente:")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Codice fiscale")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Dati anagrafici")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Codice fiscale del coobbligato")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Codice identificativo")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Sezione erario e altro")), style = "Normal") |>
            body_add_fpar(fpar(ftext("Codice ufficio e Codice atto [non compilare]")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Tipo")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Elementi identificativi [codice CIG]")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Codice")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Anno di riferimento")), style = "Elenco punto") |>
            body_add_fpar(fpar(ftext("Importi a debito versati")), style = "Elenco punto")

          print(doc, target = paste0(pre.nome.file, "5.6 Comprova imposta di bollo.docx"))

          cat("
        Documento '", pre.nome.file, "5.6 Comprova imposta di bollo.docx' generato e salvato in ", pat)
        }
    }
  }
  # Genera Comunicazione CIG ----
  com_cig <- function(){
    doc <- doc.com.cig |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("All'"), ftext("Unità Relazioni con il Pubblico", fpt.b)), style = "Destinatario") |>
      body_add_fpar(fpar(ftext("del CNR")), style = "Destinatario 2") |>
      body_add_par("", style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("OGGETTO", fpt.bi), ftext(": ", fpt.i), ftext("procedura per l’affidamento diretto, ai sensi dell’art. 50, comma 1, lettera b del D.lgs. n. 36/2023, ", fpt.i),
                         ftext(della.fornitura, fpt.i), ftext(" di “", fpt.i),
                         ftext(Prodotto, fpt.i),
                         ftext("”, Riferimento Spazio su sito URP: ", fpt.i),
                         ftext(Pagina.web, fpt.bi),
                         ftext(" - ", fpt.i), ftext("COMUNICAZIONE CIG", fpt.bi)), style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("Con riferimento alla procedura in oggetto si comunica che il CIG associato è il seguente: "), ftext(CIG, fpt.b), ftext(".")), style = "Normal") |>
      body_add_fpar(fpar(ftext("Si ringrazia per la cortese collaborazione.")), style = "Normal") |>
      body_add_fpar(fpar(ftext("Cordiali saluti.")), style = "Normal") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("Il Responsabile Unico del Progetto")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(Dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    print(doc, target = paste0(pre.nome.file, "6 Comunicazione CIG.docx"))
    cat("

    Documento '", pre.nome.file, "6 Comunicazione CIG.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, RUP, Pagina.web, CIG)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
    }
  }

  # Genera AI ----
  ## AI ----
  ai <- function(){
    doc <- doc.ai |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("ATTO ISTRUTTORIO")), style = "heading 1", pos = "on") |>
      body_add_fpar(fpar(ftext("Affidamento diretto, ai sensi dell’art. 50 del D.Lgs. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” (CIG "),
                         ftext(CIG),
                         ftext(CUI1),
                         ftext(", "), ftext(Pagina.web),
                         ftext("), nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(ordine.trattativa.scelta),
                         ftext(", ordine CNR-IPSP-"),
                         ftext(sede),
                         ftext(" N° "),
                         ftext(ordine),
                         ftext(y),
                         ftext(".")), style = "Oggetto maiuscoletto") |>
      body_add_fpar(fpar(ftext("IL RESPONSABILE UNICO DEL PROGETTO")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 6 novembre 2012, n. 190 recante “Disposizioni per la prevenzione e la repressione della corruzione e dell'illegalità nella pubblica amministrazione” pubblicata sulla Gazzetta Ufficiale n. 265 del 13/11/2012;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 14 marzo 2013, n. 33 recante “Riordino della disciplina riguardante il diritto di accesso civico e gli obblighi di pubblicità, trasparenza e diffusione di informazioni da parte delle pubbliche amministrazioni” pubblicato sulla Gazzetta Ufficiale n. 80 del 05/04/2013 e successive modifiche introdotte dal d.lgs. 25 maggio 2016 n. 97;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il D.lgs. 31 marzo 2023, n. 36 rubricato “Codice dei Contratti Pubblici in attuazione dell’articolo 1 della legge 21 giugno 2022, n. 78, recante delega al Governo in materia di contratti pubblici”, pubblicato sul Supplemento Ordinario n. 12 della GU n. 77 del 31 marzo 2023 (nel seguito per brevità “Codice”);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 50, comma 1, lettera b) del Codice, il quale consente, per affidamenti di contratti di servizi e forniture, ivi compresi i servizi di ingegneria e architettura e l'attività di progettazione di importo inferiore a euro 140.000,00, di procedere ad affidamento diretto, anche senza consultazione di più operatori economici;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento relativo all’affidamento diretto "), ftext(della.fornitura), ftext(" di cui all’oggetto, prot. "),
                         ftext(Prot..DaC),
                         ftext(" all’operatore economico "),
                         ftext(Fornitore),
                         ftext(", con sede legale in "),
                         ftext(Fornitore..Sede),
                         ftext(", C.F./P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", con il quale è "),
                         ftext(nominato),
                         ftext(" "),
                         ftext(il.dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(" quale Responsabile Unico del Progetto ai sensi dell’art. 15 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA ", fpt.b),
                         ftext(ordine.trattativa.scelta.pres)), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATI", fpt.b), ftext(" altresì i principi previsti dall’art. 57 del d.lgs. 36/2023 tra i quali le clausole sociali volte a garantire le pari opportunità generazionali, di genere e di inclusione lavorativa per le persone con disabilità o svantaggiate, la stabilità occupazionale del personale impiegato;")), style = "Normal")

    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 52, comma 1 del Codice, il quale dispone che, nelle procedure di affidamento di cui all’art. 50, comma 1, lett. b) di importo inferiore a 40.000 euro, gli operatori economici attestano con dichiarazione sostitutiva di atto di notorietà il possesso dei requisiti di partecipazione e di qualificazione richiesti e che le stazioni appaltanti procedono alla risoluzione del contratto qualora a seguito delle verifiche non sia confermato il possesso dei requisiti generali dichiarati;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato ha sottoscritto la dichiarazione sostitutiva attestante il possesso dei requisiti di ordine generale previsti dal Codice ai sensi dell’art. 52 del Codice, archiviata con prot. ")), style = "Normal") |>
        #                   ftext(Prot..DocOE), ftext(";")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che la Stazione appaltante verificherà, previo sorteggio di un campione individuato con modalità predeterminata, le dichiarazioni degli operatori economici affidatari;")), style = "Normal")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato ha sottoscritto la dichiarazione sostitutiva attestante il possesso dei requisiti di ordine generale previsti dal Codice ai sensi dell’art. 52 del Codice e il DGUE ai fini dell’avvio delle verifiche ai sensi dell’art. 94, 95, 96, 97, 98 e 100 del d.lgs. n. 36/2023 e successive modifiche ed integrazioni;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che le verifiche effettuate ai sensi dell’art. 94, 95, 96, 97, 98 e 100 del d.lgs. n. 36/2023 non hanno rilevato cause ostative nei confronti dell’operatore economico individuato;")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTI", fpt.b), ftext(" gli atti della procedura in argomento ed accertata la regolarità degli stessi in relazione alla normativa ed ai regolamenti vigenti;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VALUTATO", fpt.b), ftext(" il principio del risultato;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("AI FINI DELL’ISTRUTTORIA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("Dichiara:")), style = "Normal") |>
      body_add_fpar(fpar(ftext("Che il procedimento di acquisto risulta condotto in conformità alle disposizioni di legge e ai regolamenti vigenti in materia;")), style = "Elenco liv1")

    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Nulla osta all’emissione della lettera d’ordine purché munita di apposita clausola risolutiva in caso di accertamento della carenza dei requisiti di ordine generale.")), style = "Elenco liv1")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("Nulla osta al perfezionamento della lettera d’ordine/contratto con l’Operatore Economico individuato.")), style = "Elenco liv1")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar("Il Responsabile Unico del Progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(Dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_reach(doc, "CAMPO.DATA")
    e <- e$officer_cursor$which +1
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }

    e <- cursor_reach(doc, "CAMPO.DATA")
    e <- e$officer_cursor$which +1
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }

    ## Dich. Ass. RUP ----
    doc <- cursor_reach(doc, "SEZIONE.DICH.ASS.RICH.") |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50 DEL D.LGS. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" DI “"),
                         ftext(Prodotto),
                         ftext("” (CIG "),
                         ftext(CIG),
                         ftext(CUI1),
                         ftext(", "), ftext(Pagina.web),
                         ftext("), nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(ordine.trattativa.scelta),
                         ftext(", ordine CNR-IPSP-"),
                         ftext(sede),
                         ftext(" N° "),
                         ftext(ordine),
                         ftext(y)), style = "Maiuscolo") |>
      body_add_fpar(fpar(ftext("AUTODICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.rup), ftext(" "), ftext(RUP, fpt.b), ftext(", "),
                         ftext(nato.rup), ftext(" "), ftext(RUP..Luogo.di.nascita), ftext(", il "),
                         ftext(RUP..Data.di.nascita), ftext(", codice fiscale "), ftext(RUP..Codice.fiscale), ftext(", ")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in relazione all’incarico di Responsabile Unico del Progetto per l’affidamento "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” (CIG "),
                         ftext(CIG),
                         ftext(CUI1),
                         ftext(", "), ftext(Pagina.web),
                         ftext(")"),
                         ftext(ordine.trattativa.scelta),
                         ftext(", ordine CNR-IPSP-"),
                         ftext(sede),
                         ftext(" N° "),
                         ftext(ordine),
                         ftext(y),
                         ftext(", all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext(")"),
                         ftext(", nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                         ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di non trovarsi, rispetto al ruolo ricoperto ed alle funzioni svolte, in alcuna delle situazioni di conflitto di interessi, anche potenziale, di cui all’art. 16 del D.lgs. n. 36/2023, né nelle ipotesi previste dall’art. 35-bis, del D.lgs. n. 165/2001, tali da ledere l’imparzialità e l’immagine dell’agire dell’amministrazione;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("di impegnarsi a comunicare qualsiasi conflitto d’interesse che possa insorgere durante il presente affidamento o nella fase esecutiva del contratto;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("di impegnarsi ad astenersi prontamente dalla prosecuzione dell’affidamento diretto nel caso emerga un conflitto d’interesse;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("DICHIARA ALTRESÌ")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di aver preso piena cognizione del D.P.R. 16 aprile 2013, n. 62 e delle norme in esso contenute, nonché del Codice di comportamento dei dipendenti del Consiglio Nazionale delle Ricerche adottato con delibera del Consiglio di Amministrazione n° 137/2017;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("SI IMPEGNA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("a non utilizzare a fini privati le informazioni di cui dispone in ragione del ruolo ricoperto, a non divulgarle al di fuori dei casi consentiti e ad evitare situazioni e comportamenti che possano ostacolare il corretto adempimento della funzione sopra descritta;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("a comunicare tempestivamente eventuali variazioni del contenuto della presente dichiarazione e a rendere, se del caso, una nuova dichiarazione sostitutiva.")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("La presente dichiarazione è resa ai sensi e per gli effetti dell’art. 6-bis Legge 241/1990, degli artt. 6 e 7 del D.P.R. 16 aprile 2013, n. 62, dell’art. 53, comma 14, del D. Lgs. n° 165/2001, dell’art. 15, comma 1, lettera c) del D. Lgs. n° 33/2013 e dell’art. 20 del D. Lgs. n° 39/2013.")), style = "Normal") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da)), style = "Normal") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar("Il Responsabile Unico del Progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(Dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- cursor_reach(doc, "SEZIONE.DICH.ASS.RESP.")
    b <- b$officer_cursor$which
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc <- cursor_reach(doc, "SEZIONE.DICH.ASS.RESP.")
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    doc <- cursor_end(doc)
    doc <- cursor_backward(doc)
    doc <- cursor_backward(doc)
    doc <- cursor_backward(doc)
    doc <- body_remove(doc)
    doc <- body_remove(doc)

    print(doc, target = paste0(pre.nome.file, "7 Atto istruttorio.docx"))
    cat("

    Documento '", pre.nome.file, "7 Atto istruttorio.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, CIG, Progetto, Prot..DaC, Fornitore, Fornitore..Sede, Fornitore..P.IVA, RUP, RUP..Luogo.di.nascita, RUP..Data.di.nascita, RUP..Codice.fiscale, Pagina.web)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    I documenti sono stati generati, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente i documenti generati: i dati mancanti sono indicati con '__________'.
    **********************")
    }
  }

  # Genera Lettere d'ordine ----
  ldo <- function(){
    if(file.exists("Elenco prodotti.xlsx")=="FALSE"){
      cat("

    Premere INVIO per caricare il file Excel con l'elenco dei prodotti
        ")
      inpt <- readline()
      pr <- read.xlsx(utils::choose.files(default = "*.xlsx"))
    }else{
      pr <- read.xlsx("Elenco prodotti.xlsx")
    }

    Imponibile.ldo <- colnames(pr)[7]
    IVA.ldo <- pr[1,7]
    Importo.ldo <- pr[2,7]
    Imponibile.ldo.txt <- paste("€", format(as.numeric(Imponibile.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    IVA.ldo.txt <- paste("€", format(as.numeric(IVA.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    Importo.ldo.txt <- paste("€", format(as.numeric(Importo.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))

    pr <- pr[,1:5]
    colnames(pr) <- c("Quantità", "Descrizione", "Costo unitario senza IVA", "Importo senza IVA", "Inv./Cons.")
    pr <- subset(pr, !is.na(pr$Quantità))
    pr$`Inv./Cons.`[which(is.na(pr$`Inv./Cons.`))] <- ""
    pr$`Costo unitario senza IVA` <- paste("€", format(as.numeric(pr$`Costo unitario senza IVA`), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    pr$`Importo senza IVA` <- paste("€", format(as.numeric(pr$`Importo senza IVA`), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    prt <- pr[,-5]
    colnames(prt) <- c("Quantità", "Descrizione", "Costo unitario", "Importo")

    ## Inglese
    prt.en <- prt
    colnames(prt.en) <- c("Amount", "Description", "Unit cost", "Total")
    Prot..DaC.en <- sub("del", "of", Prot..DaC)

    doc <- doc.ldo
    b <- cursor_begin(doc)
    b <- b$officer_cursor$which
    e <- cursor_reach(doc, "CAMPO.CUP.LDO.IT")
    e <- e$officer_cursor$which -2
    doc <- cursor_begin(doc)
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }

    doc <- doc |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("LETTERA D’ORDINE CNR-IPSP-"), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
      body_add_par("") |>
      cursor_reach("CAMPO.CUP.LDO.IT") |>
      body_replace_all_text("CAMPO.CUP.LDO.IT", CUP2, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.CIG") |>
      body_replace_all_text("CAMPO.CIG", CIG, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.CUI") |>
      body_replace_all_text("CAMPO.CUI", CUI2, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.RUP") |>
      body_replace_all_text("CAMPO.RUP", RUP, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.OFFERTA.LDO") |>
      body_replace_all_text("CAMPO.OFFERTA.LDO", Preventivo.fornitore, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DAC.LDO") |>
      body_replace_all_text("CAMPO.DAC.LDO", Prot..DaC, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.RDO1") |>
      body_replace_all_text("CAMPO.RDO1", ordine.trattativa.scelta.ldo1, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.RDO2") |>
      body_replace_all_text("CAMPO.RDO2", as.character(ordine.trattativa.scelta.ldo2), only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.WEB") |>
      body_replace_all_text("CAMPO.WEB", Pagina.web, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.FORNITORE") |>
      body_replace_all_text("CAMPO.FORNITORE", Fornitore, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.SEDE") |>
      body_replace_all_text("CAMPO.SEDE", Fornitore..Sede, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.PIVA") |>
      body_replace_all_text("CAMPO.PIVA", as.character(Fornitore..P.IVA), only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.PEC") |>
      body_replace_all_text("CAMPO.PEC", Fornitore..PEC, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.EMAIL") |>
      body_replace_all_text("CAMPO.EMAIL", Fornitore..E.mail, only_at_cursor = TRUE) |>
      body_add_par("") |>
      body_add_par("") |>
      body_add_table(prt, style = "Tabella LdO", pos = "on") |>
      cursor_reach("CAMPO.IMPONIBILE") |>
      body_replace_all_text("CAMPO.IMPONIBILE", Imponibile.ldo.txt, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.ALIQUOTA") |>
      body_replace_all_text("CAMPO.ALIQUOTA", paste0("IVA (", Aliquota.IVA, ")"), only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.IVA") |>
      body_replace_all_text("CAMPO.IVA", IVA.ldo.txt, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.IMPORTO") |>
      body_replace_all_text("CAMPO.IMPORTO", Importo.ldo.txt, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.CONSEGNA") |>
      body_replace_all_text("CAMPO.CONSEGNA", Richiedente..Luogo.di.consegna, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.FATTURAZIONE") |>
      body_replace_all_text("CAMPO.FATTURAZIONE", fatturazione, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.FIRMA.LDO.IT") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2", pos = "on") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2") |>
      body_add_break() |>

      body_add_par("CONDIZIONI GENERALI D'AQCUISTO", style = "heading 1") |>
      body_add_fpar(fpar(ftext("1. Ambito di applicazione", fpt.b), ftext(": le presenti condizioni generali di acquisto hanno la finalità di regolare in modo uniforme i rapporti contrattuali con i fornitori dai quali il CNR acquista beni e/o servizi in applicazione delle norme di legge e di regolamento. Le condizioni di vendita del fornitore non saranno in nessun caso applicabili ai rapporti contrattuali con il CNR, anche se fossero state richiamate in qualsiasi documento proveniente dal fornitore stesso.")), style = "Riquadro paragrafo") |>
      body_add_fpar(fpar(ftext("2. Resa", fpt.b), ftext(": franco destino.")), style = "Riquadro paragrafo") |>
      body_add_fpar(fpar(ftext("3. Durata", fpt.b), ftext(": "), ftext(fornitura.consegnata), ftext(" entro 30 giorni naturali e consecutivi decorrenti dalla data di sottoscrizione del presente contratto presso il luogo indicato nella pagina precedente.")), style = "Riquadro paragrafo") |>
      body_add_fpar(fpar(ftext("4. Fatturazione", fpt.b), ftext(": la fattura, redatta secondo la normativa vigente, dovrà riportare, pena il rifiuto della stessa, il numero d'ordine (corrispondente al numero di registrazione al protocollo), il CIG e il CUP.")), style = "Riquadro paragrafo") |>
      body_add_fpar(fpar(ftext("5. Pagamento", fpt.b), ftext(": il pagamento sarà effettuato entro 30 gg. a partire dalla data del certificato di regolare esecuzione.")), style = "Riquadro paragrafo") |>
      body_add_fpar(fpar(ftext("6. Penali", fpt.b), ftext(": per ogni giorno naturale e consecutivo di ritardo rispetto ai termini previsti per l’esecuzione dell’appalto di cui all’art.8, si applicherà una penale pari all’1‰ (uno per mille) dell’importo contrattuale, al netto dell’IVA e dell’eventuale costo relativo alla sicurezza sui luoghi di lavoro derivante dai rischi di natura interferenziale. Per i soli contratti di forniture, nel caso in cui la prima verifica di conformità della fornitura abbia esito sfavorevole non si applicano le penali; qualora tuttavia l’Aggiudicatario non renda nuovamente la fornitura disponibile per la verifica di conformità entro i 20 (venti) giorni naturali e consecutivi successivi al primo esito sfavorevole, ovvero la verifica di conformità risulti nuovamente negativa, si applicherà la penale sopra richiamata per ogni giorno solare di ritardo. Nell’ipotesi in cui l’importo delle penali applicabili superi l’importo pari al 20% (venti per cento) dell’importo contrattuale, al netto dell’IVA e dell’eventuale costo relativo alla sicurezza sui luoghi di lavoro derivante dai rischi di natura interferenziale, l’Ente risolverà il contratto in danno all’Aggiudicatario, salvo il diritto al risarcimento dell’eventuale ulteriore danno patito.")), style = "Riquadro paragrafo") |>
      body_add_fpar(fpar(ftext("7. Tracciabilità dei flussi finanziari", fpt.b), ftext(": il fornitore assume tutti gli obblighi di tracciabilità dei flussi finanziari di cui all’art. 3 della L. 136/2010 e s.m.i. Il mancato utilizzo del bonifico bancario o postale ovvero degli altri strumenti di incasso o pagamento idonei a consentire la piena tracciabilità delle operazioni costituisce motivo di risoluzione unilaterale del contratto. Il fornitore si impegna a consentire all’Amministrazione la verifica di cui al c. 9 art. 3 della legge 136/2010 e s.m.i. e a dare immediata comunicazione all'Amministrazione ed alla Prefettura-UTG della provincia ove ha sede l'Amministrazione della notizia dell’inadempimento della propria controparte (subappaltatore/subcontraente) agli obblighi di tracciabilità finanziaria.")), style = "Riquadro paragrafo")

    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("8. Clausola risolutiva espressa", fpt.b), ftext(": l’ordine è emesso in applicazione delle disposizioni contenute all’art. 52, commi 1 e 2 del d.lgs 36/2023. Il CNR ha diritto di risolvere il contratto/ordine in caso di accertamento della carenza dei requisiti di partecipazione. Per la risoluzione del contratto trovano applicazione l’art. 122 del d.lgs. 36/2023, nonché gli articoli 1453 e ss. del Codice Civile. Il CNR darà formale comunicazione della risoluzione al fornitore, con divieto di procedere al pagamento dei corrispettivi, se non nei limiti delle prestazioni già eseguite.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("9. Foro competente", fpt.b), ftext(": per ogni controversia sarà competente in via esclusiva il Tribunale di Roma.")), style = "Riquadro paragrafo")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("8. Foro competente", fpt.b), ftext(": per ogni controversia sarà competente in via esclusiva il Tribunale di Roma.")), style = "Riquadro paragrafo")
    }

    doc <- doc |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("La presente lettera d’ordine, perfezionata mediante scambio di corrispondenza commerciale, è sottoscritta da ciascuna Parte, anche mediante sovrascrizione, con firma digitale valida alla data di apposizione della stessa e a norma di legge, ed è successivamente scambiata tra le parti via PEC. Pertanto, l’imposta di registro sarà dovuta in caso d’uso ai sensi del D.P.R 131/1986.")), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar("Per accettazione", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2")

    b <- cursor_reach(doc, "NORMATIVA DI RIFERIMENTO")
    b <- doc$officer_cursor$which + 15
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc <- cursor_reach(doc, "NORMATIVA DI RIFERIMENTO")
    doc <- cursor_forward(doc)
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    doc <- cursor_end(doc)
    doc <- body_remove(doc)
    doc <- cursor_backward(doc)
    doc <- body_remove(doc)

    if(Fornitore..Nazione=="Italiana"){
      b <- cursor_reach(doc, "CAMPO.INIZIO.LDO.EN")
      b <- doc$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which +10
      doc <- cursor_reach(doc, "CAMPO.FIRMA.LDO.EN")
      for(i in 1:(e-b)){
        doc <- body_remove(doc)
      }
      doc <- cursor_end(doc)
      doc <- body_remove(doc)
      doc <- cursor_backward(doc)
      doc <- body_remove(doc)
    }else{
      doc <- doc |>
        cursor_reach("CAMPO.INIZIO.LDO.EN") |>
        body_add_fpar(fpar(ftext("PURCHASE ORDER CNR-IPSP-"), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
        body_add_par("") |>
        cursor_reach("CAMPO.CUP.LDO.EN") |>
        body_replace_all_text("CAMPO.CUP.LDO.EN", CUP2, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.CIG") |>
        body_replace_all_text("CAMPO.CIG", CIG, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.CUI") |>
        body_replace_all_text("CAMPO.CUI", CUI2, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.RUP") |>
        body_replace_all_text("CAMPO.RUP", RUP, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.QUOTATION") |>
        body_replace_all_text("CAMPO.QUOTATION", Preventivo.fornitore, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.DAC.LDO") |>
        body_replace_all_text("CAMPO.DAC.LDO", Prot..DaC.en, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.RDO1") |>
        body_replace_all_text("CAMPO.RDO1", ordine.trattativa.scelta.ldo1, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.RDO2") |>
        body_replace_all_text("CAMPO.RDO2", as.character(ordine.trattativa.scelta.ldo2), only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.WEB") |>
        body_replace_all_text("CAMPO.WEB", Pagina.web, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.FORNITORE") |>
        body_replace_all_text("CAMPO.FORNITORE", Fornitore, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.SEDE") |>
        body_replace_all_text("CAMPO.SEDE", Fornitore..Sede, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.PIVA") |>
        body_replace_all_text("CAMPO.PIVA", as.character(Fornitore..P.IVA), only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.PEC") |>
        body_replace_all_text("CAMPO.PEC", Fornitore..PEC, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.EMAIL") |>
        body_replace_all_text("CAMPO.EMAIL", Fornitore..E.mail, only_at_cursor = TRUE) |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_table(prt.en, style = "Tabella LdO", pos = "on") |>
        cursor_reach("CAMPO.IMPONIBILE") |>
        body_replace_all_text("CAMPO.IMPONIBILE", Importo.senza.IVA, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.ALIQUOTA") |>
        body_replace_all_text("CAMPO.ALIQUOTA", paste0("VAT (", Aliquota.IVA, ")"), only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.IVA") |>
        body_replace_all_text("CAMPO.IVA", IVA, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.IMPORTO") |>
        body_replace_all_text("CAMPO.IMPORTO", Importo.con.IVA, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.CONSEGNA") |>
        body_replace_all_text("CAMPO.CONSEGNA", Richiedente..Luogo.di.consegna, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.FATTURAZIONE") |>
        body_replace_all_text("CAMPO.FATTURAZIONE", fatturazione, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.FIRMA.LDO.EN") |>
        body_add_fpar(fpar(ftext("The Responsible")), style = "Firma 2", pos = "on") |>
        body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2") |>
        body_add_break() |>

        body_add_par("GENERAL PURCHASE CONDITION", style = "heading 1") |>
        body_add_fpar(fpar(ftext("1. Scope of application", fpt.b), ftext(": These general conditions of purchase are intended to uniformly regulate contractual relationships with suppliers from whom CNR purchases goods and/or services in application of the laws and regulations. The supplier's conditions of sale will in no case be applicable to contractual relationships with CNR, even if they were referred to in any document originating from the supplier itself.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("2. Delivery", fpt.b), ftext(": to the destination.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("3. Duration", fpt.b), ftext(": "), ftext(" the order must be delivered within 30 consecutive calendar days from the date of signing this contract at the location indicated on the previous page.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("4. Invoice", fpt.b), ftext(": the invoice, drawn up in accordance with current legislation, must include, under penalty of rejection, the order number (corresponding to the protocol registration number), the CIG and the CUP.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("5. Payment", fpt.b), ftext(": payment will be made within 30 days from the date of the certificate of proper execution.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("6. Penalties", fpt.b), ftext(": for each natural and consecutive day of delay with respect to the terms provided for the execution of the contract referred to in art. 8, a penalty equal to 1‰ (one per thousand) of the contractual amount will be applied, net of VAT and any costs relating to safety in the workplace arising from risks of an interfering nature. For supply contracts only, in the event that the first conformity check of the supply has an unfavorable outcome, the penalties will not apply; however, if the Successful Bidder does not make the supply available again for the conformity check within 20 (twenty) natural and consecutive days following the first unfavorable outcome, or the conformity check is again negative, the penalty referred to above will be applied for each calendar day of delay. In the event that the amount of the applicable penalties exceeds the amount equal to 20% (twenty percent) of the contractual amount, net of VAT and any costs relating to safety in the workplace arising from interference risks, the Entity will terminate the contract to the detriment of the Successful Bidder, without prejudice to the right to compensation for any further damage suffered.")), style = "Riquadro paragrafo") |>
        body_add_fpar(fpar(ftext("7. Traceability of financial flows", fpt.b), ftext(": the supplier assumes all obligations of traceability of financial flows pursuant to art. 3 of Law 136/2010 and subsequent amendments. Failure to use bank or postal transfers or other collection or payment instruments suitable for allowing full traceability of transactions constitutes grounds for unilateral termination of the contract. The supplier undertakes to allow the Administration to carry out the verification pursuant to paragraph 9 of art. 3 of Law 136/2010 and subsequent amendments and to immediately notify the Administration and the Prefecture-UTG of the province where the Administration is based of the news of the failure of its counterpart (subcontractor/subcontractor) to comply with the obligations of financial traceability.")), style = "Riquadro paragrafo")

      if(Importo.senza.IVA.num<40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("8. Express termination clause", fpt.b), ftext(": the order is issued in application of the provisions contained in art. 52, paragraphs 1 and 2 of Legislative Decree 36/2023. The CNR has the right to terminate the contract/order in the event of a lack of participation requirements being ascertained. For the termination of the contract, art. 122 of Legislative Decree 36/2023, as well as articles 1453 et seq. of the Civil Code, apply. The CNR will formally communicate the termination to the supplier, with a ban on proceeding with the payment of the fees, except within the limits of the services already performed.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("9. Competent court", fpt.b), ftext(": the Court of Rome will have exclusive jurisdiction over any dispute.")), style = "Riquadro paragrafo")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("8. Competent court", fpt.b), ftext(": the Court of Rome will have exclusive jurisdiction over any dispute.")), style = "Riquadro paragrafo")
      }

      doc <- doc |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext("This order letter, perfected through the exchange of commercial correspondence, is signed by each Party, also by overwriting, with a digital signature valid on the date of affixing thereof and in accordance with the law, and is subsequently exchanged between the parties via PEC. Therefore, the registration tax will be due in case of use pursuant to Presidential Decree 131/1986.")), style = "Normal") |>
        body_add_par("") |>
        body_add_fpar(fpar("Signature for acceptance", run_footnote(x=block_list(fpar(ftext(" The declarant must sign with a qualified digital signature or attach a photocopy of a valid identity document (art. 38 of Presidential Decree no. 445/2000 and subsequent amendments).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2")
    }

    print(doc, target = paste0(pre.nome.file, "8 Lettera ordine.docx"))

    cat("

    Documento '", pre.nome.file, "8 Lettera ordine.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, CIG, RUP, RDO, Fornitore, Fornitore..Sede, Fornitore..P.IVA, Fornitore..PEC, Fornitore..E.mail, Importo.senza.IVA, Aliquota.IVA, Richiedente..Luogo.di.consegna, Pagina.web)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    Il documento è stato generato, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente il documento generato: i dati mancanti sono indicati con '__________'.
    **********************")
    }
  }

  # Genera Dich. Prestazione resa ----
  dic_pres <- function(){
    if(file.exists("Elenco prodotti.xlsx")=="FALSE"){
      cat("

    Premere INVIO per caricare il file Excel con l'elenco dei prodotti
        ")
      inpt <- readline()
      pr <- read.xlsx(utils::choose.files(default = "*.xlsx"))
    }else{
      pr <- read.xlsx("Elenco prodotti.xlsx")
    }

    Imponibile.ldo <- colnames(pr)[7]
    IVA.ldo <- pr[1,7]
    Importo.ldo <- pr[2,7]
    Imponibile.ldo.txt <- paste("€", format(as.numeric(Imponibile.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    IVA.ldo.txt <- paste("€", format(as.numeric(IVA.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    Importo.ldo.txt <- paste("€", format(as.numeric(Importo.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))


    doc <- doc.dic.pres |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_par("DICHIARAZIONE DI PRESTAZIONE RESA", style = "heading 1", pos = "on") |>
      body_add_par("Il Responsabile Unico del Progetto (RUP)", style = "heading 1") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il “Regolamento per le acquisizioni in economia di beni e servizi” pubblicato sulla Gazzetta Ufficiale dell’8 giugno 2013 n. 133;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento relativo all’affidamento diretto prot. "),
                         ftext(Prot..DaC), ftext(" per l'acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” (CIG "),
                         ftext(CIG),
                         ftext(CUI1),
                         ftext(", "), ftext(Pagina.web),
                         ftext("), nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(";")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTA ", fpt.b),
                         ftext(ordine.trattativa.scelta.pres)), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la lettera d’ordine CNR-IPSP-"), ftext(sede),
                         ftext(" N° "), ftext(ordine), ftext(y),
                         ftext(" di "), ftext(Importo.con.IVA),
                         ftext(" IVA inclusa emessa nei confronti dell'operatore economico "),
                         #ftext(Prot..lettera.ordine),
                         ftext(Fornitore), ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext("; codice terzo SIGLA "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(");")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il documento di trasporto;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di aver svolto la procedura secondo la normativa vigente;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext(materiale.conforme)), style = "Elenco punto") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(sede1), ftext(", __/__"), ftext(y)), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Il Responsabile Unico del Progetto (RUP)")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(Dott.rup), ftext(" "), ftext(RUP), ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    print(doc, target = paste0(pre.nome.file, "9 Dichiarazione prestazione resa.docx"))
    #cat("\014")
    #cat(rep("\n", 20))
    #cat("\014")
    cat("

    Documento '", pre.nome.file, "9 Dichiarazione prestazione resa.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Importo.con.IVA, Fornitore, Fornitore..P.IVA, Fornitore..Codice.terzo.SIGLA, Pagina.web)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    Il documento è stato generato, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente il documento generato: i dati mancanti sono indicati con '__________'.
    *********************")
    }
  }

  # Genera Provv. Liquidazione ----
  provv_liq <- function(){
    doc <- doc.prov.liq |>
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)

    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }

    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_par("PROVVEDIMENTO DI LIQUIDAZIONE E PAGAMENTO", style = "heading 1", pos = "on") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "heading 1") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Organizzazione e Funzionamento del CNR emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 01/08/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il DP CNR n. 0025034 in data 4 maggio 2005 concernente il Regolamento di amministrazione, contabilità e finanza del Consiglio Nazionale delle Ricerche e in particolare l’art. 29 “Liquidazione” e l’art. 30 “Titoli di pagamento”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Provvedimento del Presidente del CNR n. 02 del 11/01/2019 di modifica e sostituzione dell’Atto Costitutivo dell’IPSP;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Provvedimento del Presidente del CNR 26/2022 di modifica e sostituzione dell’Atto Costitutivo dell’IPSP;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 10.000,000 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento relativo all’affidamento diretto prot. "),
                         ftext(Prot..DaC), ftext(" per l'acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” ("),
                         ftext(Pagina.web),
                         ftext("), nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la lettera d’ordine CNR-IPSP-"), ftext(sede),
                         ftext(" N° "), ftext(ordine), ftext(y),
                         ftext(" di "), ftext(Importo.con.IVA),
                         #ftext(" IVA inclusa (prot. "),
                         #ftext(Prot..lettera.ordine),
                         ftext(" IVA inclusa;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento di impegno:")), style = "Normal") |>
      body_add_fpar(fpar(ftext("Fornitore: "), ftext(Fornitore), ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext("; codice terzo SIGLA "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(");")), style = "Elenco punto")

    if(CUP!=trattini){
      doc <- doc |>
        body_add_fpar(fpar(ftext("CUP: "), ftext(CUP), ftext(";")), style = "Elenco punto")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("CIG: "), ftext(CIG), ftext(";")), style = "Elenco punto")

    if(CUI!=trattini){
      doc <- doc |>
        body_add_fpar(fpar(ftext("CUI: "), ftext(CUI), ftext(";")), style = "Elenco punto")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("Impegno N° "),
                         ftext(N..impegno.di.spesa),
                         ftext(" di "),
                         ftext(Importo.con.IVA),
                         ftext(", GAE "),
                         ftext(GAE),
                         ftext(", voce di spesa "),
                         ftext(Voce.di.spesa),
                         ftext(", C/R _____, natura _____;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("Repertorio Contratti n°_____; DURC scadenza: "),
                         ftext(Fornitore..DURC.scadenza),
                         ftext(";")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VALUTATO", fpt.b),
                         ftext(" di aver ottemperato agli obblighi previsti dalla Legge 136/2010 “Tracciabilità dei flussi finanziari”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la fattura della ditta "),
                         ftext(Fornitore),
                         ftext(" N° _____ del _____ di "), ftext(Importo.con.IVA),
                         ftext(", scadenza _____. SDI registrata in attività _____;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b),
                         ftext(" che le prestazioni rese sono state regolarmente eseguite, come attestato nella dichiarazione di prestazione resa prot. "),
                         ftext(Prot..prestazione.resa),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DISPONE")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("la liquidazione della succitata fattura ed autorizza il responsabile amministrativo all’emissione del relativo mandato di pagamento su IBAN: "),
      ftext(Fornitore..IBAN), ftext(".")), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2") |>
      body_end_section_continuous()

    b <- doc$officer_cursor$which +1
    e <- cursor_end(doc)
    e <- e$officer_cursor$which
    doc$officer_cursor$which <- b
    for(i in 1:(e-b)){
      doc <- body_remove(doc)
    }
    doc <- body_remove(doc)
    doc <- body_remove(doc)
    print(doc, target = paste0(pre.nome.file, "10 Provv. liquidazione.docx"))
    cat("

    Documento '", pre.nome.file, "10 Provv. liquidazione.docx' generato e salvato in ", pat)

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prot..DaC, Prodotto, Fornitore, Fornitore..P.IVA, Fornitore..Codice.terzo.SIGLA, CIG, N..impegno.di.spesa, Importo.con.IVA, GAE, Voce.di.spesa, Pagina.web)
    manca <- as.data.frame(t(manca))
    colnames(manca) <- "val"
    manca$var <- rownames(manca)
    rownames(manca) <- NULL
    manca <- subset(manca, manca$val==trattini)
    len <- length(manca$val)
    if(len>0){
      manca <- manca$var
      manca <- paste0(manca, ",")
      manca[len] <- sub(",$", "\\.", manca[len])
      cat("
    ***** ATTENZIONE *****
    Il documento è stato  generato, ma i seguenti dati risultano mancanti:", manca)
      cat("
    Si consiglia di leggere e controllare attentamente il documento generato: i dati mancanti sono indicati con '__________'.
    **********************")
    }
  }

  # Input ----
  answ <- function(){
    cat("\014")
    #cat(rep("\n", 20))
    cat("\014")
    cat("

      ***************************
      *** BENVENUTI in AppOst ***
      ***************************

      Che documento vuoi generare?
      1: RAS, con eventuale avviso pubblico, Richiesta pagina web
      2: Provvedimento d'impegno, Decisione a contrattare
      3: Comunicazione CIG, Documenti dell'Operatore Economico, Atto istruttorio, Lettera d'ordine, Dichiarazione di prestazione resa, Provvedimento di liquidazione

      ")
    inpt <- readline()
    if(inpt==1){ras();pag()}
    if(inpt==2){provv_imp();dac()}
    if(inpt==3){com_cig();docoe();ai();ldo();dic_pres();provv_liq()}
    #if(inpt==4){provv_liq()}
    # if(inpt==5){
    #   # drive_deauth()
    #   # drive_user()
    #   # elenco.prodotti <- drive_get(as_id("1Hqjc3fruTBy04u_ULwua1Cegtbs7ndEe"))
    #   # drive_download(elenco.prodotti, overwrite = TRUE)
    # download.file("https://raw.githubusercontent.com/giovabubi/appost/main/models/Elenco%20prodotti.xlsx", destfile = "Elenco prodotti.xlsx", method = "curl")
    # cat("\014")
    # #cat(rep("\n", 20))
    # cat("\014")
    # cat("
    #   #
    #   # Documento 'Elenco prodotti.xlsx' generato e salvato in ", pat)
    # }

    cat("

    Vuoi generare altri documenti di quest'ordine?
      1: Sì
      2: No
      ")
    inpt2 <- readline()
    if(inpt2==1){if(interactive()) answ()}
  }
  cat("\014")
  if(interactive()) answ()
}
