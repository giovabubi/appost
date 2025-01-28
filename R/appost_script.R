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
  #ordine <- "AGRITECH-FI 01"
  #ordine <- 193
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
  }else if(file.exists("Ordini_2025.csv")=="TRUE"){
    ordini <- read.csv("Ordini_2025.csv", na.strings = "")
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

  # if(file.exists("Elenco prodotti.xlsx")=="FALSE"){
  #   cat("
  #
  #   Premere INVIO per caricare il file Excel con l'elenco dei prodotti
  #       ")
  #   inpt <- readline()
  #   file.elenco.prodotti <- utils::choose.files(default = "*.xlsx")
  #   n <- stringr::str_locate_all(file.elenco.prodotti, "\\\\")
  #   m <- max(n[[1]])
  #   n <- paste0(".{", m, "}")
  #   file.elenco.prodotti2 <- sub(n, "", file.elenco.prodotti)
  # }else{
  #   file.elenco.prodotti <- "Elenco prodotti.xlsx"
  # }

  if(!require(dplyr)) install.packages("dplyr")
  ordini <- dplyr::rename(ordini,
                          Prodotto=Descrizione.beni.servizi.lavori,
                          RDO=N..RDO.MePA,
                          sede=Sede)
  colnames(ordini)[3] <- "Data"
  ordini$Fornitore..P.IVA <- as.character(ordini$Fornitore..P.IVA)
  ordini$CPV <- NULL
  ordini$CPV <- as.character(ordini$CPV..CPV)
  ordini$Importo.senza.IVA.num <- sub(",(..)$", "_\\1", ordini$Importo.senza.IVA)
  ordini$Importo.senza.IVA.num <- gsub("\\.", "", ordini$Importo.senza.IVA.num)
  ordini$Importo.senza.IVA.num <- gsub("_", ".", ordini$Importo.senza.IVA.num)
  ordini$Importo.senza.IVA.num <- as.numeric(ordini$Importo.senza.IVA.num)

  sc <- subset(ordini, ordini$Ordine.N.==ordine)
  
  anno <- grep("_2024", sc$Ordine.N.)
  if(length(anno==1)){
    y <- "/2024"
    y2 <- 2024
    sc$Ordine.N. <- sub("_2024", "", sc$Ordine.N.)
    ordine <- sub("_2024", "", ordine)
  }else{
    y <- "/2025"
    y2 <- 2025
  }

  sc$Aliquota.IVA.num <- as.numeric(ifelse(sc$Aliquota.IVA=='22%', 0.22,
                                           ifelse(sc$Aliquota.IVA=='10%', 0.1,
                                                  ifelse(sc$Aliquota.IVA=='4%', 0.04, 0))))
  sc$IVA <- sc$Importo.senza.IVA.num * sc$Aliquota.IVA.num
  sc$Importo.con.IVA <- sc$Importo.senza.IVA.num + sc$IVA
  sc$Importo.senza.IVA <- paste("€", format(sc$Importo.senza.IVA.num, format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
  sc$IVA <- paste("€", format(sc$IVA, format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
  sc$Importo.con.IVA <- paste("€", format(sc$Importo.con.IVA, format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
  sc$Importo.senza.IVA2 <- paste0(sub("€ ", "", sc$Importo.senza.IVA), " €")
  sc$IVA2 <- paste0(sub("€ ", "", sc$IVA), " €")
  sc$Importo.con.IVA2 <- paste0(sub("€ ", "", sc$Importo.con.IVA), " €")

  ## Installa e carica pacchetti ---
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
    al.RSS <- "Al responsabile della sede secondaria di Bari"
    firma.RSS <- "Il responsabile della sede secondaria di Bari dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, via G. Amendola 122/D, 70126 Bari, Italia."
    nomina.RSS <- "3903 dell'8/1/2025 di nomina del dott. Giovanni Nicola Bubici quale Responsabile della Sede Secondaria di Bari dell’IPSP per il periodo dall’8/1/2025 al 31/12/2025"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Nicola Centorame quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "Il responsabile amministrativo della sede secondaria di Bari dell'IPSP"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a Foggia il giorno 11/11/1977, codice fiscale BBCGNN77S11D643H,"
    RSS.dich <- "responsabile della sede secondaria di Bari dell'IPSP"
  }else if(sede=='TO'){
    sede1 <- 'Torino'
    sede2 <- 'Sede Secondaria di Torino'
    RSS <- 'Dott. Stefano Ghignone'
    RSS.email <- 'stefano.ghignone@cnr.it'
    RAMM <- "Dott.ssa Lucia Allione"
    RAMM.email <- 'lucia.allione@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Torino"
    firma.RSS <- "Il responsabile della sede secondaria di Torino dell'IPSP"
    in.qualita.RSS <- "responsabile della sede secondaria di Torino dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, viale Mattioli, 25, 10125 Torino, Italia."
    nomina.RSS <- "3906 dell'8/1/2025 di nomina del dott. Stefano Ghignone quale Responsabile della Sede Secondaria di Torino dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della dott.ssa Lucia Allione quale Responsabile Amministrativo della Sede Secondaria di Torino dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La responsabile amministrativa della sede secondaria di Torino dell'IPSP"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a __________ il __________, codice fiscale __________,"
    RSS.dich <- " responsabile della sede secondaria di Torino "
  }else if(sede=='NA'){
    sede1 <- 'Portici'
    sede2 <- 'Sede Secondaria di Portici'
    RSS <- 'Dott.ssa Michelina Ruocco'
    RSS.email <- 'michelina.ruocco@cnr.it'
    RAMM <- 'Dott. Ettore Magaldi'
    RAMM.email <- 'ettore.magaldi@ipsp.cnr.it'
    al.RSS <- "Alla responsabile della sede secondaria di Portici"
    firma.RSS <- "La responsabile della sede secondaria di Portici dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, piazzale Enrico Fermi, 1, 80055 Portici (NA), Italia."
    nomina.RSS <- "3907 dell'8/1/2025 di nomina della dott.ssa Michelina Ruocco quale Responsabile della Sede Secondaria di Portici dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Ettore Magaldi quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "Il responsabile amministrativo"
    sottoscritto.rss <- "La sottoscritta "
    nato.rss <- " nata a __________ il __________, codice fiscale __________,"
    RSS.dich <- " responsabile della sede secondaria di Portici dell'IPSP "
  }else if(sede=='FI'){
    sede1 <- 'Sesto Fiorentino'
    sede2 <- 'Sede Secondaria di Sesto Fiorentino'
    RSS <- "Dott. Nicola Luchi"
    RSS.email <- "nicola.luchi@ipsp.cnr.it"
    RAMM <- "Sig.ra Francesca Pesciolini"
    RAMM.email <- 'francesca.pesciolini@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Sesto Fiorentino"
    firma.RSS <- "Il responsabile della sede secondaria di Sesto Fiorentino dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, via Madonna del Piano, 10, 50019 Sesto F.no (FI), Italia."
    nomina.RSS <- "3904 dell'8/1/2025 di nomina del dott. Nicola Luchi quale Responsabile della Sede Secondaria di Sesto Fiorentino dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RAMM <- "146220 del 2/5/2024 di nomina della sig.ra Francesca Pesciolini quale Responsabile Amministrativo della Sede Secondaria di Sesto Fiorentino dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La responsabile amministrativa"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a __________ il __________, codice fiscale __________,"
    RSS.dich <- " responsabile della sede secondaria di Sesto Fiorentino dell'IPSP "
  }else if(sede=='PD'){
    sede1 <- 'Legnaro'
    sede2 <- 'Sede Secondaria di Legnaro'
    RSS <- "Dott.ssa Laura Scarabel"
    RSS.email <- "laura.scarabel@ipsp.cnr.it"
    RAMM <- "Dott.ssa Lucia Allione"
    RAMM.email <- 'lucia.allione@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Legnaro"
    firma.RSS <- "Il responsabile della sede secondaria di Legnaro dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, viale dell’Università, 16, 35020 Legnaro (PD), Italia."
    nomina.RSS <- "3905 dell'8/1/2025 di nomina della dott.ssa Laura Scarabel quale Responsabile della Sede Secondaria di Legnaro dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della dott.ssa Lucia Allione quale Responsabile Amministrativo della Sede Secondaria di Legnaro dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La responsabile amministrativa"
    sottoscritto.rss <- "La sottoscritta "
    nato.rss <- " nata a __________ il __________, codice fiscale __________,"
    RSS.dich <- " responsabile della sede secondaria di Legnaro dell'IPSP "
  }else if(sede=='TOsi'){
    sede1 <- 'Torino'
    sede2 <- 'Sede Istituzionale'
    RSS <- 'Dott. Francesco Di Serio'
    RSS.email <- 'francesco.diserio@cnr.it'
    RAMM <- 'Dott. Josè Saporita'
    RAMM.email <- 'jose.saporita@ipsp.cnr.it'
    al.RSS <- "Al direttore dell'IPSP"
    firma.RSS <- "Il direttore dell'IPSP"
    in.qualita.RSS <- "direttore dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, Strada delle Cacce, 73, 10135 Torino, Italia."
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Josè Saporita quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La segretaria amministrativa"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a __________ il __________, codice fiscale __________,"
    RSS.dich <- " direttore dell'IPSP "
  }

  if(Scelta.fornitore=='Avviso pubblico'){
    preventivo.individuato <- paste0("stato individuato per motivazioni tecnico-scientifiche e di economicità tra i preventivi di ",
                                     Prot..preventivi.avviso,
                                     " pervenuti in seguito all'avviso pubblico prot. ",
                                     Prot..avviso.pubblico,
                                     ";")
  }else if(Scelta.fornitore=='Più preventivi'){
    preventivo.individuato <- "individuato a seguito di indagine informale di mercato effettuata su MePA, mercato libero e/o cataloghi accessibili in rete con esito allegato alla richiesta medesima e ritenuto in grado di assicurare la fornitura o la prestazione secondo i tempi e le modalità indicati dall’amministrazione, garantendo le migliori condizioni economiche e tecnico-qualitative;"
  }else{
    preventivo.individuato <- "allegato alla richiesta medesima e ritenuto in grado di assicurare la fornitura o la prestazione secondo i tempi e le modalità indicati dall’amministrazione, garantendo le migliori condizioni economiche e tecnico-qualitative;"
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
  if(Supporto.RUP..Sesso=='M'){
    Dott.sup <- 'Dott.'
    dott.sup <- 'dott.'
    il.dott.sup <- 'il dott.'
    al.sup <- 'al'
    dal.sup <- 'dal dott.'
    Nato.sup <- 'Nato a'
    nato.sup <- 'nato a'
    assegna <- 'assegnatario'
    sottoscritto.sup <- 'Il sottoscritto'
  }else{
    Dott.sup <- 'Dott.ssa'
    dott.sup <- 'dott.ssa'
    il.dott.sup <- 'la dott.ssa'
    al.sup <- 'alla'
    dal.sup <- 'dalla dott.ssa'
    Nato.sup <- 'Nata a'
    nato.sup <- 'nata a'
    assegna <- 'assegnataria'
    sottoscritto.sup <- 'La sottoscritta'
  }

  da <- as.character(Sys.Date())
  # y <- sub("(....)-(..)-(..)",  "/\\1", da)
  # y2 <- sub("(....)-(..)-(..)",  "\\1", da)
  da <- sub("(....)-(..)-(..)",  "\\3/\\2/\\1", da)

  if(length(anno==1)){
  }else{
    anno <- grep("\\/2024$", sc$Prot..RAS)
  }
  
  #if(ordine>40 | ordine<1){
  if(length(anno==1)){
    y <- "/2024"
    y2 <- 2024
  }else{
    y <- "/2025"
    y2 <- 2025
  }
  
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
  all.OE <- paste0("all'operatore economico ", Fornitore, " (P.IVA ", Fornitore..P.IVA, ")")
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
                    ", ordine ", sede, " N° ", ordine, y, ".", sep=""))

  #pre.nome.file <- paste0("Ordine ", sede, " ", ordine, "_", y2, " - ")
  pre.nome.file <- paste0("Ordine ", ordine, "_", y2, " - ")
  
  ### Durata affidamento
  dat1 <- as.Date(sub(".*del ", "", Prot..DaC), "%d/%m/%Y")
  dat2 <- as.Date(sub(".*del ", "", Prot..prestazione.resa), "%d/%m/%Y")
  gg <- as.numeric(difftime(dat2, dat1, units = "days"))
  dat1 <- format(dat1, "%d/%m/%Y")
  dat2 <- format(dat2, "%d/%m/%Y")
  durata.affidamento <- paste0(gg, " giorni, dal ", dat1, " al ", dat2, " (decisione a contrattare prot. n. ", Prot..DaC, "; dichiarazione prestazione resa prot. n. ", Prot..prestazione.resa, ")")
  
  ### PNRR ----
  if(PNRR!="No"){
    lnk <- "https://raw.githubusercontent.com/giovabubi/appost/main/models/PNRR/"
  }else{
    lnk <- "https://raw.githubusercontent.com/giovabubi/appost/main/models/"
    dicitura.fattura <- trattini
  }
  if(PNRR=="Agritech Spoke 3" | PNRR=="Agritech Spoke 8" | PNRR=="Agritech Spoke 4" | PNRR=="onFoods Spoke 4"){
    finanziamento <- "PNRR"
  }else if(PNRR=="DIVINGRAFT" | PNRR=="ARES" | PNRR=="MINACROP" | PNRR=="MONTANA" | PNRR=="SpecFor" | PNRR=="Mimic" | PNRR=="StreeTLAMP" | PNRR=="Fore-VOC"){
    finanziamento <- "PRIN 2022"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 104 del 2/2/2022 di emanazione del bando per i progetti di rilevante interesse nazionale (PRIN) 2022, nell’ambito del piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2, investimento 1.1;"
    mis.com.inv.esteso <- "piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.1 “progetti di ricerca di significativo interesse nazionale (PRIN)”, finanziamento dell'Unione europea - NextGeneration EU, decreto direttoriale MUR n. 104 del 2/2/2022"
    investimento <- "Investimento 1.1 “progetti di ricerca di rilevante interesse nazionale (PRIN)”"
  }else if(PNRR=="CIRCUFUN" | PNRR=="KNOWS" | PNRR=="PEP-HERB" | PNRR=="NEUROPLANT" | PNRR=="SAVEASH"){
    finanziamento <- "PRIN 2022 PNRR"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 1409 del 14/9/2022 di emanazione del bando per i progetti di rilevante interesse nazionale (PRIN) 2022, nell’ambito del piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2, investimento 1.1;"
    mis.com.inv.esteso <- "piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.1 “progetti di ricerca di significativo interesse nazionale (PRIN)”, finanziamento dell'Unione europea - NextGeneration EU, decreto direttoriale MUR n. 104 del 2/2/2022"
    investimento <- "Investimento 1.1 “progetti di ricerca di rilevante interesse nazionale (PRIN)”"
  }
  
  if(PNRR=="Agritech Spoke 3" | PNRR=="Agritech Spoke 8" | PNRR=="Agritech Spoke 4"){
    Progetto.int <- 'piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.4 “potenziamento strutture di ricerca e creazione di campioni nazionali di R&S su alcune key enabling technologies”, finanziato dall’Unione europea, NextGenerationEU, decreto direttoriale MUR n. 1032 del 17/06/2022 - progetto “centro nazionale tecnologie dell’agricoltura” (Agritech), spoke 3, codice identificativo CN00000022 - CUP B83C22002840001'
    codice.progetto <- "CN00000022"
    CUP2 <- "B83C22002840001"
    decreto.concessione <- "1032 del 17/6/2022"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 3138 del 16/12/2021 di emanazione di un “Avviso pubblico per la presentazione di Proposte di intervento per il Potenziamento di strutture di ricerca e creazione di campioni nazionali di R&S su alcune Key Enabling Technologies” da finanziare nell’ambito del Piano Nazionale di Ripresa e Resilienza, missione 4, componente 2, investimento 1.4 “Potenziamento strutture di ricerca e creazione di campioni nazionali di R&S su alcune Key Enabling Technologies”, finanziato dall’Unione europea – NextGenerationEU;"
    decreto.pnrr <- " il Decreto Direttoriale MUR n. 1032 del 17/6/2022 rettificato con Decreto Direttoriale n. 3175 del 18/12/2021, registrato alla Corte dei Conti l’8/7/2022 al n. 1826 (di seguito, “Decreto di Concessione del Finanziamento”) col quale è stata ammessa a finanziamento la domanda di agevolazione presentata dal Centro Nazionale “National Research Centre for Agricultural Technologies”, tematica “Tecnologie dell’Agricoltura (Agritech)”, contrassegnata dal codice identificativo CN00000022, per la realizzazione del Programma di Ricerca dal titolo “NationalResearch Centre for Agricultural Technologies”, CUP B83C22002840001;"
    dicitura.fattura <- "PNRR AGRITECH - Codice progetto MUR: CN00000022"
    investimento <- "Investimento 1.4 “potenziamento strutture di ricerca e creazione di campioni nazionali di R&S su alcune key enabling technologies”"
    intervento <- "Centro Nazionale di Ricerca per le Tecnologie dell’Agricoltura (Agritech), codice progetto CN00000022"
    attuatore <- "Centro Nazionale per le Tecnologie dell’Agricoltura “National Research Centre for Agricultural Technologies” (Agritech)"
    avvio <- "Avvio: 1/9/2022; Conclusione: 31/8/2025"
    costo.totale <- "346.342.467,00 €"
    costo.ammesso <- "320.070.095,50 €"
    logo <- "logo_agritech.png"
  }
  if(PNRR=="Agritech Spoke 4"){
    Progetto.int <- sub("spoke 3", "spoke 4", Progetto.int)
  }
  if(PNRR=="Agritech Spoke 8"){
    Progetto.int <- sub("spoke 3", "spoke 8", Progetto.int)
    dicitura.fattura <- "Finanziamento Unione Europea NextGenerationEU progetto PNRR AGRITECH Spoke8 M4.C2.I1.4 - Codice progetto MUR: CN00000022"
  }
  if(PNRR=="onFoods Spoke 4"){
    Progetto <- "onFoods Spoke 4"
    Progetto.int <- 'piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.3 “partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base”, finanziato dall’Unione europea, NextGenerationEU, decreto direttoriale MUR n. 1550 dell’11/10/2022 - progetto “Research and innovation network on food and nutrition sustainability, safety and security - working on foods” (ON Foods), spoke 4, codice identificativo PE00000003, CUP B83C22004790001'
    codice.progetto <- "PE00000003"    
    CUP2 <- "B83C22004790001"
    decreto.concessione <- "1550 dell’11/10/2022"
    avviso.pnrr <- " l'avviso pubblico n. 341 del 15/03/2022 per la presentazione di proposte di intervento per la creazione di “Partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base” nell’ambito del piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.3 “partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base”, finanziato dall’Unione europea, NextGenerationEU;"
    decreto.pnrr <- " il Decreto Direttoriale MUR n. 1243 del 02-08-2022, con il quale il quale il MUR approva gli esiti delle valutazioni delle proposte progettuali presentate nell’ambito del sopra citato Avviso Pubblico con il quale viene ammesso a finanziamento il Progetto PNRR - Codice progetto PE00000003, PE10 - Modelli per un’alimentazione sostenibile, “ON Foods” - Research and innovation network on food and nutrition Sustainability, Safety and Security – Working ON Foods;"
    decreto.pnrr2 <- " il Decreto Direttoriale di concessione n. 1550 dell’11 ottobre 2022, ha ammesso al finanziamento, il Partenariato Esteso denominato ON Foods” - Research and innovation network on food and nutrition Sustainability, Safety and Security – Working ON Foods” corredato dagli allegati A-B-C-D-E;"
    decreto.pnrr3 <- " tutti gli allegati al decreto di Concessione suddetto e in particolare l’allegato A con la scheda tecnica della proposta progettuale;"
    decreto.pnrr4 <- " che il CNR ha acquisito il CUP B83C22004790001;"
    decreto.pnrr5 <- " la comunicazione MUR AOODGRIC.REGISTRO UFFICIALE Prot. N. 0017196 del 14/09/2023 che ha approvato la nuova proposta di Allegato B “Piano dei costi e delle agevolazioni” e Allegato C “Cronoprogramma di attuazione e Piano dei Pagamenti” che pertanto sostituiscono quelli precedenti e che costituiscono parte integrante del decreto direttoriale di concessione;"
    decreto.pnrr6 <- " l’articolo 27, ai commi 2 e 3, del DL 24 febbraio 2023, n. 13, convertito con modificazioni dalla Legge 21 aprile 2023, n. 41 prevede per le università statali, per gli enti pubblici di ricerca di cui all’articolo 1, comma 1, del d.lgs. n. 218/2016 e per le istituzioni statali dell’alta formazione artistica, musicale e coreutica, l’utilizzo dei propri sistemi interni di gestione e controllo al fine di assicurare il corretto impiego delle risorse finanziarie assegnate nell’ambito delle misure del PNRR/PNC, nonché il raggiungimento degli obiettivi in conformità alle disposizioni generali, sia normative che amministrative, di contabilità pubblica;"
    decreto.pnrr7 <- " la circolare esplicativa adottata dal MUR prot. n. 3739 del 22.5.2023 recante 'Modalità di rendicontazione del Decreto Legge 24 febbraio 2023, n. 13, convertito con modificazioni dalla Legge 21 aprile 2023, n. 41' recepisce la suddetta modalità semplificata e fornisce le indicazioni al riguardo in relazione al perimetro di applicazione della norma;"
    decreto.pnrr8 <- " l’Accordo HUB_SPOKE_Affiliati Atto giuridico del Progetto ON Foods approvato con Delibera n. 214/2023 dal Consiglio di Amministrazione del CNR nella seduta del 20 giugno 2023 perfezionato da tutti partecipanti;"
    decreto.pnrr9 <- " che il CNR partecipa al Progetto ON Foods nel ruolo sia di Spoke Leader (Spoke 2) che di Affiliato agli Spoke 3 (Università di Bari), Spoke 4 (Università di Milano), Spoke 5 (Università di Napoli) e Spoke 6 (Università di Pavia) e che il contributo a favore del CNR, pari a complessivi € 9.210.000,00 di cui € 4.600.000 al Mezzogiorno, è finanziato al 100%;"
    decreto.pnrr10 <- " che il CNR partecipa allo Spoke 3 nel ruolo di Affiliato all’Università di Bari e che ai sensi dell’art. 5 dell’Accordo per la regolamentazione dei rapporti tra Hub ed i Soggetti Realizzatori (Spoke e Affiliati), ciascun Affiliato è tenuto a provvedere alla rendicontazione delle proprie spese, che saranno approvate dallo Spoke di riferimento;"
    decreto.pnrr11 <- " che nello Spoke 3 il CNR (IPSP) Istituto per la Protezione Sostenibile delle Piante, con Sede istituzionale a Torino partecipa in qualità di capofila CNR nei confronti degli altri Istituti CNR partecipanti (CNR-ISPA, CNR-IBBR, CNR-ISAFOM, CNR-ISB e CNR-ISA) e che il budget totale della partecipazione CNR allo Spoke 3 è pari a € 1.850.000,20;"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    investimento <- "Investimento 1.3 “partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base”"
    intervento <- "On Foods – Research and Innovation Network On Food and Nutrition Sustainability, Safety and Security – Working On Foods, codice progetto PE0000003"
    attuatore <- "Fondazione OnFoods"
    avvio <- "Avvio: 1/11/2022; Conclusione: 31/10/2025"
    costo.totale <- "115.303.750,00 €"
    costo.ammesso <- "114.500.000,00 €"
    logo <- "logo_onfoods.jpg"
  }
  if(PNRR=="DIVINGRAFT"){
    titolo <- "dissection of molecular mechanisms underlying tolerance to virus and viroid infection in grafted tomato plants"
    codice.progetto <- "2022BZW9PF"
    CUP2 <- "B53D23017480006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche (CNR), Istituto per la Protezione Sostenibile delle Piante (IPSP)"
    avvio <- "Avvio: 12/10/2023; Conclusione: 11/10/2025"
    costo.totale <- "264.979,00 €, di cui al CNR-IPSP 133.549,00 €"
    costo.ammesso <- "206.154,00 €, di cui al CNR-IPSP 103.900,00 €"
    logo <- "logo_divingraft.jpg"
  }
  if(PNRR=="CIRCUFUN"){
    titolo <- "assessment of biodiversity and functional roles of infectious circular viroid-like RNAs in fungi from different ecological niches"
    codice.progetto <- "P2022XX55J"
    CUP2 <- "B53D23023750001"
    decreto.concessione <- "1180 del 27/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche (CNR), Istituto per la Protezione Sostenibile delle Piante (IPSP)"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "237.299,00 €, di cui al CNR-IPSP 114.690,00 €"
    costo.ammesso <- "237.299,00 €, di cui al CNR-IPSP 114.690,00 €"
    logo <- "logo_circufun.png"
  }
  if(PNRR=="ARES"){
    titolo <- "apricot genomics and transcriptomics to unravel the genetic bases of resistance to Sharka and the plant/virus interaction"
    codice.progetto <- "2022F79TR4_LS9_PRIN2022"
    CUP2 <- "B53D23017580006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Bologna, DiSTAL (capofila)"
    avvio <- "Avvio: 12/10/2023; Conclusione: 11/10/2025"
    costo.totale <- "206.194,00 € (rimodulato, ad intero carico MUR), di cui a Unità CNR (IPSP-BA) 66.897,00 €"
    costo.ammesso <- "206.194,00 € (rimodulato, ad intero carico MUR), di cui a Unità CNR (IPSP-BA) 66.897,00 €"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="KNOWS"){
    titolo <- "Generating KNOWledge on insect-pathogen-agroecosystem interaction for a Sustainable Xylella fastidiosa control"
    codice.progetto <- "P2022LJ5TM_LS9_PRIN2022PNRR"
    CUP2 <- "B53D2303218 0001"
    decreto.concessione <- "1377 dell'1/9/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Bari Aldo Moro (capofila)"
    avvio <- "Avvio: 1/12/2023; Conclusione: 30/11/2025"
    costo.totale <- "224.963,00 €, di cui a Unità CNR (IPSP-BA) 36.924,00 €"
    costo.ammesso <- "224.963,00 €, di cui a Unità CNR (IPSP-BA) 36.924,00 €"
    logo <- "logo_knows.jpeg"
  }
  if(PNRR=="MINACROP"){
    titolo <- "The dark side of MIcro- and NAanoplastics in the soil: impact on CROP physiology and pathogen resistance"
    codice.progetto <- "2022LF3SZE_LS8_PRIN2022"
    CUP2 <- "B53D23012120006"
    decreto.concessione <- "1015 del 7/7/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Ivan BACCELLI")
    attuatore <- "Università degli Studi di Firenze (capofila)"
    avvio <- "Avvio: 5/10/2023; Conclusione: 4/10/2025"
    costo.totale <- "239.149,00 €, di cui a Unità CNR (IPSP-FI) 60.109,00 €"
    costo.ammesso <- "239.149,00 €, di cui a Unità CNR (IPSP-FI) 60.109,00 €"
    logo <- "logo_minacrop.png"
  }
  if(PNRR=="MONTANA"){
    titolo <- "Ulmus glabra protection in Italian peninsula"
    codice.progetto <- "2022SFNMYC_LS8_PRIN2022"
    CUP2 <- "B53D23012340006"
    decreto.concessione <- "1015 del 7/7/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Alessia Lucia PEPORI")
    attuatore <- "CNR, Isttuto per la Protezione Sostenibile delle Piante"
    avvio <- "Avvio: 5/10/2023; Conclusione: 4/10/2025"
    costo.totale <- "234.924,00 €, di cui a Unità CNR (IPSP-FI) 173.305,00 €"
    costo.ammesso <- "185.671,00 €, di cui a Unità CNR (IPSP-FI) 133.251,00 €"
    logo <- "logo_montana.png"
  }
  if(PNRR=="SpecFor"){
    titolo <- "spectroscopic detection of forest damage: investigating new Italian holm oak declines from leaf to landscape level"
    codice.progetto <- "20223AR37M"
    CUP2 <- "B53D23017080006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. dott.ssa BARBERINI")
    attuatore <- "Università di Pisa"
    avvio <- "Avvio: 14/10/2023; Conclusione: 13/10/2025"
    costo.totale <- "320.101,00 €, di cui a Unità CNR (IPSP-FI) 103.163,00 €"
    costo.ammesso <- "267.724,00 €, di cui a Unità CNR (IPSP-FI) 85.316,00 €"
    logo <- "logo_specfor.jpg"
  }
  if(PNRR=="Mimic"){
    titolo <- "mimic the interplay of phytohormones and Biogenic Volatile Organic Compounds (BVOC) by genome editing approaches to boost rice meristem development and yield"
    codice.progetto <- "2022NKSSKM_LS2_PRIN2022"
    CUP2 <- "B53D23008290006"
    decreto.concessione <- "970 del 30/6/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Federico BRILLI")
    attuatore <- "Università degli Studi di Milano"
    avvio <- "Avvio: 30/9/2023; Conclusione: 29/9/2025"
    costo.totale <- "299.674,00 €, di cui a Unità CNR (IPSP-FI) 154.132,00 €"
    costo.ammesso <- "248.429,00 €, di cui a Unità CNR (IPSP-FI) 100.500,00 €"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="StreeTLAMP"){
    titolo <- "new streetlamp solution to reduce the impact of urban light pollution on tree and lichen species"
    codice.progetto <- "20222YF92Y_LS9 _PRIN2022"
    CUP2 <- "B53D23017060006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Cecilia BRUNETTI")
    attuatore <- "Università di Pisa"
    avvio <- "Avvio: 18/9/2023; Conclusione: 17/6/2025"
    costo.totale <- "233.000,00 €, di cui a Unità CNR (IPSP-FI) 58.000,00 €"
    costo.ammesso <- "193.024,00 €, di cui a Unità CNR (IPSP-FI) 49.200,00 €"
    logo <- "logo_streetlamp.jpg"
  }
  if(PNRR=="Fore-VOC"){
    titolo <- "plants talk, but do they listen? Unveiling plant responses to incoming (foreign) volatile organic compounds"
    codice.progetto <- "2022ZYCCJJ_LS8_PRIN2022"
    CUP2 <- "B53D23012480006"
    decreto.concessione <- "1015 del 7/7/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Susanna POLLASTRI")
    attuatore <- "Università di Napoli"
    avvio <- "Avvio: 5/10/2023; Conclusione: 4/10/2025"
    costo.totale <- "313.279,00 €, di cui a Unità CNR (IPSP-FI) 119.146,00 €"
    costo.ammesso <- "257.828,00 €, di cui a Unità CNR (IPSP-FI) 98.056,00 €"
    logo <- "logo_forevoc.png"
  }
  if(PNRR=="PEP-HERB"){
    titolo <- "developing PEPtide molecules targeting the plant immune system to fight HERBicide-resistant weeds"
    codice.progetto <- "P2022NEJ8K_LS9_PRIN2022PNRR"
    CUP2 <- "B53D23032230001"
    decreto.concessione <- "1377 dell'1/9/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Ivan BACCELLI")
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "224.324,00 €, di cui a Unità CNR (IPSP-FI) 73.387,00 €"
    costo.ammesso <- "224.324,00 €, di cui a Unità CNR (IPSP-FI) 73.387,00 €"
    logo <- "logo_pepherb.jpg"
  }
  if(PNRR=="NEUROPLANT"){
    titolo <- "the development of an artificial intelligence tool to predict phytoremediation of indoor air, through NEURal netwOrks trained with measurements of pollutants removal by PLANTs and their associated microbiome at different spatial and temporal scale"
    codice.progetto <- "P2022MNX4S_LS9_PRIN2022PNRR"
    CUP2 <- "B53D23032230001"
    decreto.concessione <- "1377 dell'1/9/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Federico BRILLI")
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 1/12/2023; Conclusione: 30/11/2025"
    costo.totale <- "299.541,00 €, di cui a Unità CNR (IPSP-FI) 153.976,00 €"
    costo.ammesso <- "224.656,00 €, di cui a Unità CNR (IPSP-FI) 115.482,00 €"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="SAVEASH"){
    titolo <- "SAVE the Apennine aSh from Hymenoscyphus fraxineus"
    codice.progetto <- "P2022CHMFZ_LS9_PRIN2022PNRR"
    CUP2 <- "B53D23032100001"
    decreto.concessione <- "1377 dell'1/9/2023"
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 104/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Alberto SANTINI")
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "264.789,00 €, di cui a Unità CNR (IPSP-FI) 119.585,00 €"
    costo.ammesso <- "222.305,00 €, di cui a Unità CNR (IPSP-FI) 102.000,00 €"
    logo <- "logo_saveash.jpg"
  }
  
  dicitura.fatturazione <- paste0("Si prega di riportare in fattura le seguenti informazioni: ordine n° ", sede, " ", ordine, y, ", prot. n. _____ (si veda in alto nella pagina della lettera d'ordine), CIG ", CIG, ", CUP ", CUP, ".")
  dicitura.fatturazione.eng <- paste0("In the invoice, plese report the following information: purchase order n° ", sede, " ", ordine, y, ", prot. n. _____ (see on the top of the purchase order page), CIG ", CIG, ", CUP ", CUP, ".")
  
  if(PNRR!="No"){
    dicitura.fatturazione <- sub(".$", paste0(" e la seguente dicitura: '", dicitura.fattura, "'."), dicitura.fatturazione)
    dicitura.fatturazione.eng <- sub(".$", paste0(" and the following phrase: '", dicitura.fattura, "'."), dicitura.fatturazione.eng)
  }
  
  if(finanziamento=="PRIN 2022" | finanziamento=="PRIN 2022 PNRR"){
    Progetto <- paste(finanziamento, PNRR)
    Progetto.int <- paste0(mis.com.inv.esteso, " - progetto “", PNRR, ": ", titolo, "”, codice identificativo ", codice.progetto, ", CUP ", CUP2)
    decreto.pnrr <- paste0(" il Decreto Direttoriale MUR n. ", decreto.concessione, " (di seguito, “Decreto di Concessione del Finanziamento”) col quale è stata ammessa a finanziamento la domanda di agevolazione del progetto ", finanziamento, " “", PNRR, ": ", titolo, "”, codice identificativo ", codice.progetto, ", CUP ", CUP2, ";")
    intervento <- paste0(finanziamento, ": “", titolo, "” (", PNRR, "). Codice progetto: ", codice.progetto)
  }
  Progetto.cup <- paste0(Progetto, " (CUP ", CUP2, ")")
  if(PNRR=="No"){
    Progetto.int <- paste("progetto", Progetto.cup)
  }
  
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
                         ftext(", ORDINE "),
                         ftext(sede),
                         ftext(" "),
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
                         ftext("”, ordine "),
                         ftext(sede),
                         ftext(" "),
                         ftext(ordine),
                         ftext(y),
                         ftext(" "),
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
                           ftext(", ORDINE "),
                           ftext(sede),
                           ftext(" "),
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
                           ftext("”, ordine "),
                           ftext(sede),
                           ftext(" "),
                           ftext(ordine),
                           ftext(y),
                           ftext(" "),
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
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }
    
    doc <- doc |>
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
                         ftext(", ordine "),
                         ftext(sede),
                         ftext(" "),
                         ftext(ordine),
                         ftext(y),
                         ftext(".")), style = "Oggetto") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 31 dicembre 2009 n. 213, recante “Riordino del Consiglio Nazionale delle Ricerche in attuazione dell’articolo 1 della Legge 27 settembre 2007, n. 165“;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 25 novembre 2016 n. 218, recante “Semplificazione delle attività degli enti pubblici di ricerca ai sensi dell'articolo 13 della legge 7 agosto 2015, n. 124”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 7 agosto 1990, n. 241 recante “Nuove norme in materia di procedimento amministrativo e di diritto di accesso ai documenti amministrativi” pubblicata sulla Gazzetta Ufficiale n. 192 del 18/08/1990 e s.m.i.;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Regolamento di Organizzazione e Funzionamento del Consiglio Nazionale delle Ricerche emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 01/08/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il D.lgs. 31 marzo 2023, n. 36 rubricato “Codice dei Contratti Pubblici in attuazione dell’articolo 1 della legge 21 giugno 2022, n. 78, recante delega al Governo in materia di contratti pubblici”, pubblicato sul Supplemento Ordinario n. 12 della GU n. 77 del 31 marzo 2023 (nel seguito per brevità “Codice”);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l'art. 32 'Decisione di contrattare' del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall'art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 6 novembre 2012, n. 190 recante “Disposizioni per la prevenzione e la repressione della corruzione e dell’illegalità nella pubblica amministrazione” pubblicata sulla G.U.R.I. n. 265 del 13/11/2012;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 14 marzo 2013, n. 33 recante “Riordino della disciplina riguardante il diritto di accesso civico e gli obblighi di pubblicità, trasparenza e diffusione di informazioni da parte delle pubbliche amministrazioni” pubblicato sulla Gazzetta Ufficiale n. 80 del 05/04/2013 e successive modifiche introdotte dal d.lgs. 25 maggio 2016 n. 97;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Codice di comportamento dei dipendenti del Consiglio Nazionale delle Ricerche approvato con delibera del Consiglio di Amministrazione n° 137/2017;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il vigente Piano triennale per la prevenzione della corruzione e della trasparenza (PTPCT), adottato con delibera del Consiglio di Amministrazione del Consiglio Nazionale delle Ricerche ai sensi della legge 6 novembre 2012 n. 190;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 23 dicembre 1999 n 488 e s.m.i., recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2000)”, ed in particolare l'articolo 26;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 27 dicembre 2006 n. 296, recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2007)”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 24 dicembre 2007 n. 244 e s.m.i., recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2008)”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il decreto-legge 7 maggio 2012 n. 52, convertito dalla legge 6 luglio 2012 n. 94 recante “Disposizioni urgenti per la razionalizzazione della spesa pubblica”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il decreto-legge 6 luglio 2012 n. 95, convertito con modificazioni dalla legge 7 agosto 2012 n. 135, recante “Disposizioni urgenti per la revisione della spesa pubblica con invarianza dei servizi ai cittadini”;")), style = "Normal")
      if(Oneri.sicurezza==trattini){
        doc <- doc |>
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
                             ftext(preventivo.individuato)), style = "Normal")
      }else{
        doc <- doc |>
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
                             ftext(" oltre IVA, comprensivo di "),
                             ftext(Oneri.sicurezza),
                             ftext(" quali oneri per la sicurezza dovuti a rischi da interferenze, il cui preventivo è "),
                             ftext(preventivo.individuato)), style = "Normal")
      }
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 50, comma 1, lettera b) del Codice, il quale consente, per affidamenti di contratti di servizi e forniture, ivi compresi i servizi di ingegneria e architettura e l'attività di progettazione di importo inferiore ad euro 140.000,00, di procedere ad affidamento diretto, anche senza consultazione di più operatori economici;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(", altresì, che la scelta dell’Operatore Economico deve essere effettuata assicurando che i soggetti individuati siano in possesso di documentate esperienze pregresse idonee all’esecuzione delle prestazioni contrattuali, anche individuati tra gli iscritti in elenchi o albi istituiti dalla stazione appaltante;")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Comunicato del Presidente ANAC del 10 gennaio 2024 con cui sono state diramate indicazioni di carattere transitorio sull’applicazione delle disposizioni del codice dei contratti pubblici in materia di digitalizzazione degli affidamenti di importo inferiore a euro 5.000,00"),
      #                   ftext(" [valido fino al 30.09.2024];", fpt.i)), style = "Normal") |>
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
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il bilancio di previsione del Consiglio Nazionale delle Ricerche per l’esercizio finanziario 2025, approvato dal Consiglio di Amministrazione in data 17 dicembre 2024 con deliberazione n° 420/2024 – Verb. 511;")), style = "Normal") |>
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
      body_add_fpar(fpar(ftext("rilasciare apposita dichiarazione, rispetto al ruolo ricoperto ed alle funzioni svolte, nella quale attesti di non trovarsi in alcuna delle situazioni di conflitto di interessi, anche potenziale, di cui all’art. 16 del D.lgs. n. 36/2023;")), style = "Elenco liv2")
      
    if(Supporto.RUP!=trattini){
        doc <- doc |>
          body_add_fpar(fpar(ftext("DI INDIVIDUARE", fpt.b), ftext(" ai sensi dell’art. 15, comma 6 del Codice, "),
                             ftext(il.dott.sup), ftext(" "),
                             ftext(Supporto.RUP),
                             ftext(" in qualità di supporto al RUP;")), style = "Elenco liv1")
      }
    doc <- doc |>
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
      #body_add_par("Il Responsabile Amministrativo", style = "Firma 1") |>
      #body_add_fpar(fpar(ftext("("), ftext(RAMM), ftext(")")), style = "Firma 1") |>
      body_add_par("La segretaria amministrativa", style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(sig.ra Concetta Mottura)")), style = "Firma 1") |>
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
    #print(doc, target = paste0(pre.nome.file, "4 Decisione a contrattare per URP.docx"))

    #cat("\014")
    #cat(rep("\n", 20))
    cat("

    Documento '", pre.nome.file, "4 Decisione a contrattare.docx' generato e salvato in ", pat)

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
      headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    }
    
    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("CdR 121.001.000 IPSP ", fpt.b), ftext(sede2, fpt.b)), style = "Normal") |>
      body_add_fpar(fpar(ftext("PROVVEDIMENTO DI IMPEGNO DELLA")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("LETTERA D'ORDINE "), ftext(sede), ftext(" "), ftext(ordine), ftext(y)), style = "heading 1") |>
      body_add_fpar(fpar(firma.RSS), style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Organizzazione e Funzionamento del CNR emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 01/08/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento CNR n. 114 del 30/10/2013 (prot. n. 0065484) relativo alla costituzione dell’Istituto per la Protezione Sostenibile delle Piante con successivi provvedimenti del Presidente n. 120 del 07/10/2014 (prot. n. 72102) e n. 2 del 11/01/2019 di conferma e sostituzione del precedente atto costitutivo;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
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
      body_add_fpar(fpar(ftext("Il responsabile unico del progetto (RUP)")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
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
                               ftext(", ordine "),
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
                         ftext(", ordine "),
                         ftext(sede),
                         ftext(" "),
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
                         ftext(", ordine "),
                         ftext(sede),
                         ftext(" "),
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
                         ftext(", ordine "),
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

  # Genera Lettera d'ordine ----
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
      body_add_fpar(fpar(ftext("LETTERA D’ORDINE "), ftext(sede), ftext(" "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
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

      body_add_par("CONDIZIONI GENERALI D'ACQUISTO", style = "heading 1") |>
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
        body_add_fpar(fpar(ftext("PURCHASE ORDER "), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
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

        body_add_par("GENERAL PURCHASE CONDITIONS", style = "heading 1") |>
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
                         ftext(" la lettera d’ordine "), ftext(sede),
                         ftext(" "), ftext(ordine), ftext(y),
                         ftext(" di "), ftext(Importo.ldo.txt),
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
                         ftext(" il Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Provvedimento del Presidente del CNR n. 02 del 11/01/2019 di modifica e sostituzione dell’Atto Costitutivo dell’IPSP;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Provvedimento del Presidente del CNR 26/2022 di modifica e sostituzione dell’Atto Costitutivo dell’IPSP;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
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
                         ftext(" la lettera d’ordine "), ftext(sede),
                         ftext(" "), ftext(ordine), ftext(y),
                         ftext(" di "), ftext(Importo.ldo.txt),
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
                         ftext(" N° _____ del _____ di "), ftext(Importo.ldo.txt),
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
  # RAS PNRR ----
  ras.pnrr <- function(){
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
    
    download.file(paste(lnk, "RAS.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    if(PNRR!="No"){
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
        headers_replace_text_at_bkm(bookmark = "bookmark_headers", toupper(Progetto.int))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers", sede1)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
    }
    
    doc <- doc |>
      cursor_reach("CAMPO.DEST.RAS.SEDE") |>
      body_replace_all_text("CAMPO.DEST.RAS.SEDE", al.RSS, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DEST.RAS.RSS") |>
      body_replace_all_text("CAMPO.DEST.RAS.RSS", RSS, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DEST.RAS.EMAIL") |>
      body_replace_all_text("CAMPO.DEST.RAS.EMAIL", RSS.email, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.DELLA.FORNITURA") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("OGGETTO", fpt.b),
                         ftext(": RICHIESTA DI ACQUISTO E RELAZIONE PER L’ACQUISIZIONE "),
                         ftext(toupper(della.fornitura)),
                         ftext(" DI “"),
                         ftext(toupper(Prodotto), fpt.b),
                         ftext("”, "),
                         ftext(toupper("ordine "), fpt.b),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(toupper(ordine), fpt.b),
                         ftext(toupper(y), fpt.b),
                         ftext(", NELL'AMBITO DEL "),
                         ftext(toupper(Progetto.int)),
                         ftext(".")), style = "Normal") |>
      #body_replace_all_text("CAMPO.DELLA.FORNITURA", toupper(paste0(della.fornitura, " DI “", Prodotto, "”, ordine ",
      #                                                              sede, " N° ", ordine, y, ", NELL'AMBITO DEL ", Progetto.int, ".")), only_at_cursor = TRUE) |>
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
    
    print(doc, target = paste0(pre.nome.file, "1 RAS.docx"))
    
    cat("\014")
    cat("

    Documento '", pre.nome.file, "1 RAS.docx' generato e salvato in ", pat)
    
    ## Dich. Ass. RICH ----
    if(PNRR!="No"){
      download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.ras |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.ric), ftext(" "), ftext(dott.ric), ftext(" "), ftext(Richiedente, fpt.b), ftext(", "),
                         ftext(nato.ric), ftext(" "), ftext(Richiedente..Luogo.di.nascita), ftext(" il "),
                         ftext(Richiedente..Data.di.nascita), ftext(", codice fiscale "), ftext(Richiedente..Codice.fiscale), ftext(", ")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di "),
                         ftext("richiedente", fpt.b),
                         ftext(" l'affidamento diretto "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext(") all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                         ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
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
      body_add_fpar(fpar(ftext("")), style = "Normal")
    if(Richiedente!=Responsabile.progetto){
      doc <- doc |>
      body_add_fpar(fpar("Il richiedente l'affidamento", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.ric),
                         ftext(" "),
                         ftext(Richiedente),
                         ftext(")")), style = "Firma 2")
    }else{
      doc <- doc |>
        body_add_fpar(fpar("Il richiedente l'affidamento, responsabile del progetto e titolare dei fondi", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.ric),
                           ftext(" "),
                           ftext(Richiedente),
                           ftext(")")), style = "Firma 2")
    }
    doc <- doc |>
      body_add_break()
    print(doc, target = paste0(pre.nome.file, "4.1 Dichiarazione assenza conflitto RICH.docx"))
    
    ## Dich. Ass. RESP ----
    if(Richiedente!=Responsabile.progetto){
      if(PNRR!="No"){
        download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- doc |>
          footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
        file.remove("tmp.docx")
        file.remove(logo)
      }else{
        doc <- doc.ras |>
          headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
        }
        file.remove("tmp.docx")
      }
      
      doc <- doc |>
        cursor_begin() |>
        body_add_fpar(fpar(ftext("All’"),
                           ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
        body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext(sottoscritto.resp), ftext(" "), ftext(dott.resp), ftext(" "), ftext(Responsabile.progetto, fpt.b), ftext(", "),
                           ftext(nato.resp), ftext(" "), ftext(Responsabile.progetto..Luogo.di.nascita), ftext(", il "),
                           ftext(Responsabile.progetto..Data.di.nascita), ftext(", codice fiscale "), ftext(Responsabile.progetto..Codice.fiscale), ftext(", ")), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b),
                           ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in relazione all'affidamento diretto "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto, fpt.b),
                           ftext("”, ordine "),
                           ftext(sede, fpt.b),
                           ftext(" "),
                           ftext(ordine, fpt.b),
                           ftext(y, fpt.b),
                           ftext(" ("),
                           ftext(Pagina.web),
                           ftext(") all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), "),
                           ftext("titolare dei fondi e responsabile", fpt.b),
                           ftext(" del "),
                           ftext(Progetto.int),
                           ftext(";")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                           ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
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
        body_add_fpar(fpar("Il titolare dei fondi e responsabile del progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.resp),
                           ftext(" "),
                           ftext(Responsabile.progetto),
                           ftext(")")), style = "Firma 2") |>
        body_add_break()
      print(doc, target = paste0(pre.nome.file, "4.2 Dichiarazione assenza conflitto RESP.docx"))
    }
    
    cat("

    Dichiarazioni di assenza conflitto di interesse generate e salvate in ", pat)
    
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
  
  # RUP PNRR ----
  rup.pnrr <- function(){
    cat("\014")
    
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
    
    download.file(paste(lnk, "RUP.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))

    doc <- doc |>
      cursor_reach("CAMPO.DELLA.FORNITURA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("OGGETTO", fpt.b),
                         ftext(": NOMINA DEL RESPONSABILE UNICO DEL PROGETTO AI SENSI DELL’ART. 15 E DELL’ALLEGATO I.2 DEL DECRETO LEGISLATIVO 31 MARZO 2023 N. 36 E IMPEGNO PROVVISORIO DELLE SOMME NECESSARIE PER L’AFFIDAMENTO DIRETTO "),
                         ftext(toupper(della.fornitura)),
                         ftext(" DI “"),
                         ftext(toupper(Prodotto), fpt.b),
                         ftext("”"),
                         ftext(", ORDINE "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(toupper(Pagina.web)),
                         ftext("), NELL'AMBITO DEL "),
                         ftext(toupper(Progetto.int)),
                         ftext(".")), style = "Normal") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      cursor_reach("CAMPO.NOMINE") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")
    
    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      cursor_reach("CAMPO.DECRETO") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(avviso.pnrr)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr)), style = "Normal")
      if(PNRR=="onFoods Spoke 4"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr2)), style = "Normal") |>
          body_add_fpar(fpar(ftext("VISTI", fpt.b), ftext(decreto.pnrr3)), style = "Normal") |>
          body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr4)), style = "Normal") |>
          body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(decreto.pnrr5)), style = "Normal") |>
          body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr6)), style = "Normal") |>
          body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(decreto.pnrr7)), style = "Normal") |>
          body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr8)), style = "Normal") |>
          body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(decreto.pnrr9)), style = "Normal") |>
          body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(decreto.pnrr10)), style = "Normal") |>
          body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(decreto.pnrr11)), style = "Normal") 
      }
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la "), ftext("richiesta di acquisto prot. ", fpt.b),
                         ftext(Prot..RAS, fpt.b), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” ("),
                         ftext(Pagina.web),
                         ftext("), nell’ambito delle attività previste dal "),
                         ftext(Progetto.cup),
                         ftext(", corredata dal preventivo d'importo pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA, formulato dall'operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), "),
                         ftext(preventivo.individuato)), style = "Normal") |>
      cursor_reach("CAMPO.DISPONIBILITA") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b), ftext(" la disponibilità finanziaria per la copertura della spesa sui fondi del già richiamato progetto, allocati al GAE "),
                         ftext(GAE),
                         ftext(", voce del piano "), ftext(Voce.di.spesa), ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATA", fpt.b), ftext(", pertanto, la necessità di procedere:")), style = "Normal") |>
      body_add_fpar(fpar(ftext(" - alla nomina del responsabile unico del progetto (RUP) per la programmazione, progettazione, affidamento e l’esecuzione dell’affidamento "),
      ftext(della.fornitura), ftext(" di “"),
      ftext(Prodotto), ftext("”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext(" - all’impegno provvisorio delle somme necessarie indicate nella richiesta d’acquisto prot. n. "),
                         ftext(Prot..RAS), ftext(";")), style = "Normal") |>
      cursor_reach("CAMPO.NOMINA.RUP") |>
      body_replace_all_text("CAMPO.NOMINA.RUP", paste(il.dott.rup, RUP), only_at_cursor = TRUE) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A1", Importo.senza.IVA2) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A4", Importo.senza.IVA2) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A", Importo.senza.IVA2) |>
      body_replace_text_at_bkm(bookmark = "bookmark_B3", IVA2) |>
      body_replace_text_at_bkm(bookmark = "bookmark_B", IVA2) |>
      body_replace_text_at_bkm(bookmark = "bookmark_TOT", Importo.con.IVA2)
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_replace_text_at_bkm(bookmark = "bookmark_B1", "35")
    }
    doc <- doc |>
      cursor_reach("CAMPO.DI.IMPEGNARE") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext(Importo.con.IVA),
      ftext(" IVA inclusa sui fondi del progetto "),
      ftext(Progetto.cup),
      ftext(", voce del piano "),
      ftext(Voce.di.spesa),
      ftext(", sul GAE "),
      ftext(GAE),
      ftext(";")), style = "Elenco punto liv2")
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("€ 35,00, Voce del piano 13096 “Pubblicazione bandi di gara” sul GAE [completare] per la quota stazione appaltante della contribuzione ANAC;")), style = "Elenco punto liv2")
    }
        #   4.	(eventuale, solo nel caso di servizi diversi da quelli di natura intellettuale e di forniture con posa in opera) DI DARE ATTO che:
    #   •	ai sensi dell’art.11 del D.Lgs. 36/2023, ai dipendenti dell’O.E. affidatario dovrà essere applicato il CCNL [completare] ovvero un diverso CCNL avente le medesime tutele;
    # •	i costi della manodopera indicati nel quadro economico sopra riportato sono stati calcolati sulla base delle tariffe orarie previste per il CCNL [completare]; 
    # 5.	(eventuale) DI DEROGARE alla quota del 30% delle assunzioni necessarie di occupazione femminile e giovanile di cui all’art. 47 del decreto 77/2021 in quanto [completare indicando le motivazioni dell’eventuale deroga];
    doc <- doc |>
      body_replace_text_at_bkm(bookmark = "bookmark_prot_ras", Prot..RAS) |>
      cursor_reach("CAMPO.FIRMA") |>
      body_remove() |>
      body_add_par("Visto di regolarità contabile", style = "Firma 1") |>
      #body_add_par(resp.segr, style = "Firma 1") |>
      #body_add_fpar(fpar(ftext("("), ftext(RAMM), ftext(")")), style = "Firma 1") |>
      body_add_par("La segretaria amministrativa", style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(sig.ra Concetta Mottura)")), style = "Firma 1") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "2 Nomina RUP.docx"))
    
    cat("\014")
    cat("

    Documento '", pre.nome.file, "2 Nomina RUP.docx' generato e salvato in ", pat)
    
    ## Dich. Ass. RSS ----
    download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.rss), ftext(RSS, fpt.b), ftext(","), ftext(nato.rss)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di "),
                         ftext(RSS.dich, fpt.b),
                         ftext(" e in relazione all'affidamento diretto "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext("), all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                         ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
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
      body_add_fpar(fpar(firma.RSS, run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(RSS),
                         ftext(")")), style = "Firma 2") |>
      body_add_break()
    
    print(doc, target = paste0(pre.nome.file, "4.3 Dichiarazione assenza conflitto RSS.docx"))
    
    cat("

    Documento '", pre.nome.file, "4.3 Dichiarazione assenza conflitto RSS.docx' generato e salvato in ", pat)
    
    ## Dich. Ass. SUP ----
    if(Supporto.RUP!=trattini){
      
      doc <- read_docx("tmp.docx")
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      
      doc <- doc |>
        cursor_begin() |>
        body_add_fpar(fpar(ftext("All’"),
                           ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
        body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext(sottoscritto.sup), ftext(" "), ftext(dott.sup), ftext(" "), ftext(Supporto.RUP, fpt.b), ftext(", "), 
                           ftext(nato.sup), ftext(" "), ftext(Supporto.RUP..Luogo.di.nascita),
                           ftext(" il "), ftext(Supporto.RUP..Data.di.nascita),
                           ftext(", codice fiscale "), ftext(Supporto.RUP..Codice.fiscale), ftext(",")), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b),
                           ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di "),
                           ftext("supporto al RUP", fpt.b),
                           ftext(" nella procedura di affidamento diretto "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto, fpt.b),
                           ftext("”, ordine "),
                           ftext(sede, fpt.b),
                           ftext(" "),
                           ftext(ordine, fpt.b),
                           ftext(y, fpt.b),
                           ftext(" ("),
                           ftext(Pagina.web),
                           ftext(") all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), nell'ambito del "),
                           ftext(Progetto.int),
                           ftext(";")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                           ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
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
        body_add_fpar(fpar("Il supporto al RUP", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.sup),
                           ftext(" "),
                           ftext(Supporto.RUP),
                           ftext(")")), style = "Firma 2") |>
        body_add_break()
      
      print(doc, target = paste0(pre.nome.file, "4.5 Dichiarazione assenza conflitto SUP.docx"))
    
    cat("

    Documento '", pre.nome.file, "4.5 Dichiarazione assenza conflitto SUP.docx' generato e salvato in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Importo.senza.IVA, Voce.di.spesa, GAE, RUP, Prot..RAS, Pagina.web, RUP)
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
    file.remove("tmp.docx")
    file.remove(logo)
  }
  
  # DocOE PNRR ----
  docoe.pnrr <- function(){
    if(PNRR!="No"){
      docuOE <- c("Dich_requisiti_infra40",
                  "Dich_requisiti_over40",
                  "DPCM",
                  "Dich_tit",
                  "Dich_aus",
                  "CCNL",
                  "bollo")
      docuOE_ext <- c("3.1 Dichiarazione possesso requisiti di partecipazione e di qualificazione infra 40k",
                      "3.1 Dichiarazione possesso requisiti di qualificazione oltre 40k",
                      "3.4 Dichiarazione DPCM 187 1991",
                      "3.6 Dichiarazione titolare effettivo",
                      "3.8 Dichiarazione ausiliaria",
                      "3.9 Comprova equivalenza tutele CCNL",
                      "3.11 Comprova imposta di bollo")
      if(Importo.senza.IVA.num<40000){
        docuOE <- docuOE[c(-2,-7)]
        docuOE_ext <- docuOE_ext[c(-2,-7)]
      }else{
        docuOE <- docuOE[-1]
        docuOE_ext <- docuOE_ext[-1]
      }
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      j <- 1
      for(i in docuOE){
        download.file(paste(lnk, paste0(i, ".docx"), sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        doc <- doc |>
          footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
          body_replace_text_at_bkm(bookmark = "bookmark_body", toupper(paste0(della.fornitura, " DI “", Prodotto, "”, ordine ",
                                                                              sede, " N° ", ordine, y, " (", Pagina.web, ") NELL'AMBITO DEL ",  Progetto.int)))
        print(doc, target = paste0(pre.nome.file, paste0(docuOE_ext[j], ".docx")))
        j <- j+1
      }
    }
    
    ## CC dedicato ----
    download.file(paste(lnk, "cc_dedicato.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    print(doc, target = paste0(pre.nome.file, "3.2 Comunicazione conto corrente dedicato.docx"))
    
    ## Patto d'integrità ----
    download.file(paste(lnk, "Patto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_body", toupper(paste0(della.fornitura, " DI “", Prodotto, "”, ordine ",
                                                                          sede, " N° ", ordine, y, " (", Pagina.web, ") NELL'AMBITO DEL ",  Progetto.int))) |>
      cursor_bookmark("bookmark_OE") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("L'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (di seguito Operatore Economico) con sede legale in "),
                         ftext(Fornitore..Sede),
                         ftext(", C.F./P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", rappresentato da "),
                         ftext(Fornitore..Rappresentante.legale),
                         ftext(" in qualità di "),
                         ftext(tolower(Fornitore..Ruolo.rappresentante)),
                         ftext(",")), style = "Normal") |>
      body_replace_text_at_bkm(bookmark = "bookmark_firma", firma.RSS)
    print(doc, target = paste0(pre.nome.file, "3.3 Patto di integrità.docx"))
    
    ## CAM ----
    download.file(paste(lnk, "CAM.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
    print(doc, target = paste0(pre.nome.file, "3.10 Documentazione rispetto CAM.docx"))
    
    ## DNSH ----
    download.file(paste(lnk, "DNSH_gen.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      cursor_bookmark("bookmark_body") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("in relazione all'affidamento diretto "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext(")"),
                         ftext(", nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Elenco punto")
    print(doc, target = paste0(pre.nome.file, "3.5 Scheda DNSH generica.docx"))
    
    download.file(paste(lnk, "DNSH_app.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      cursor_bookmark("bookmark_body") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("in relazione all'affidamento "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext(")"),
                         ftext(", nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Elenco punto")
    print(doc, target = paste0(pre.nome.file, "3.5 Scheda DNSH apparecchiature.docx"))
    
    download.file(paste(lnk, "DNSH_chi.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_body", toupper(Progetto.int)) |>
      cursor_bookmark("bookmark_body") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("in relazione all'affidamento "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext(")"),
                         ftext(", nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Elenco punto")
    print(doc, target = paste0(pre.nome.file, "3.5 Scheda DNSH chimici.docx"))
    
    ## Dich.Ass. TIT ----
    download.file(paste(lnk, "Dich_conf_tit.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      cursor_bookmark("bookmark_fornitura") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("relativamente all'affidamento diretto "),
                         ftext(della.fornitura),
                         ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int)), style = "Normal")
    
    print(doc, target = paste0(pre.nome.file, "3.7 Dichiarazione assenza conflitto interesse titolare effettivo.docx"))
    
    file.remove("tmp.docx")
    file.remove(logo)
    
    cat("

    Autocertificazioni dell'operatore economico generati e salvati in ", pat)
  }
  
  # AI PNRR ----
  ai.pnrr <- function(){
    cat("\014")
 
    if(PNRR!="No"){
      download.file(paste(lnk, "Istruttoria.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.ras |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
    }
    
    doc <- doc |>
      cursor_reach("CAMPO.OGGETTO") |>
      #body_replace_all_text("CAMPO.OGGETTO", toupper(paste0("AFFIDAMENTO ", della.fornitura, " DI “", Prodotto, "”, ordine ", sede, " N° ", ordine, y, " (", Pagina.web, "), NELL'AMBITO DEL ", Progetto.int, ".")), only_at_cursor = TRUE) |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO "),
                         ftext(toupper(della.fornitura)),
                         ftext(" DI “"),
                         ftext(toupper(Prodotto), fpt.b),
                         ftext("”"),
                         ftext(", ORDINE "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(toupper(Pagina.web)),
                         ftext("), NELL'AMBITO DEL "),
                         ftext(toupper(Progetto.int)),
                         ftext(".")), style = "Normal") |>
      body_add_par("Il responsabile unico del progetto (RUP)", style = "heading 2") |>
      cursor_reach("CAMPO.PROT.RAS") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la "), ftext("richiesta di acquisto prot. ", fpt.b),
                         ftext(Prot..RAS, fpt.b), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” ("),
                         ftext(Pagina.web),
                         ftext("), nell’ambito delle attività previste dal "),
                         ftext(Progetto.cup),
                         ftext(", corredata dal preventivo d'importo pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA, formulato dall'operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), "),
                         ftext(preventivo.individuato)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il "), ftext("provvedimento prot. n. ", fpt.b),
                         ftext(Prot..provv..impegno, fpt.b),
                         ftext(", con il quale è stato nominato "),
                         ftext(il.dott.rup),
                         ftext(" "),
                         ftext(RUP, fpt.b),
                         ftext(" quale Responsabile Unico del Progetto ("),
                         ftext("RUP", fpt.b),
                         ftext(") ai sensi dell’art. 15 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(") ha presentato, attraverso la piattaforma telematica di negoziazione ("),
                         ftext("RDO ", fpt.b),
                         ftext(as.character(RDO), fpt.b),
                         ftext("), un’offerta ritenuta congrua, corredata dalle dichiarazioni sostitutive richieste, in merito al possesso dei requisiti prescritti dalla stazione appaltante, d’importo uguale o inferiore rispetto a quello stimato dalla stazione appaltante, pari a "),
                         ftext(Importo.senza.IVA, fpt.b),
                         ftext(" oltre IVA;")), style = "Normal")
    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATO altresì", fpt.b), ftext(" che sono state effettuate le verifiche, tramite l’utilizzo del sistema FVOE e degli ulteriori Enti certificatori o delle banche dati disponibili, delle dichiarazioni presentate in sede di affidamento in ordine al rispetto dei principi DNSH, alle generalità del titolare effettivo, agli obblighi assunzionali di cui all’art. 47 del decreto legge n. 77/2021, mediante acquisizione dei seguenti documenti:")), style = "Normal")
    }else{
      doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATO altresì", fpt.b), ftext(" che sono state effettuate le verifiche, tramite l’utilizzo del sistema FVOE e degli ulteriori Enti certificatori o delle banche dati disponibili, delle dichiarazioni presentate in sede di affidamento in ordine al rispetto dei principi DNSH, alle generalità del titolare effettivo, agli obblighi assunzionali di cui all’art. 47 del decreto legge n. 77/2021, nonché in ordine all’assenza delle cause di esclusione di cui agli artt. 94 e 95 del Codice ed eventuale possesso dei requisiti di cui all’art. 100 del codice, richiesti per l’esecuzione del contratto, mediante acquisizione dei seguenti documenti:")), style = "Normal")
    }
    doc <- doc |>
      body_add_fpar(fpar(ftext("(indicare le certificazioni fornite per il rispetto dei principi DNSH);")), style = "Elenco punto") |> 
      body_add_fpar(fpar(ftext("(indicare la documentazione acquisita per verificare i dati del/dei Titolare/i effettivo/i;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("(indicare la documentazione presentata per il rispetto di quanto riportato all’art. 47 del decreto legge n. 77/2021 ovvero per dimostrare l’applicabilità di eventuali deroghe allo stesso);")), style = "Elenco punto")
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |> 
        body_add_fpar(fpar(ftext("certificati generali del Casellario Giudiziale dei soggetti di cui all’art. 94, comma 3 e comma 4, del Codice dai quali non risultano a carico degli interessati elementi ostativi a contrattare con la Pubblica Amministrazione (ai sensi del comma 1 dell’art. 94 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("certificato dell’Anagrafe delle sanzioni amministrative dipendenti da reato dal quale non risultano annotazioni (ai sensi del comma 5 – lett. a) dell’art. 94 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("certificazione di ottemperanza da cui risulta che l’operatore economico è in regola con le disposizioni di legge (ai sensi del comma 5 – lett. b) dell’art. 94 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("visura camerale attestante l’inesistenza di procedure concorsuali in corso o pregresse (ai sensi del comma 5 – lett. d) dell’art. 94 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("elenco per estratto delle annotazioni iscritte nel casellario informatico tenuto dall'ANAC da cui non risultano annotazioni che comportino l’esclusione dell’operatore economico (ai sensi del comma 5 – lett. e) e f) dell’art. 94 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("esito informativo relativo alla condizione di regolarità fiscale rispetto all’assolvimento degli obblighi relativi al pagamento di imposte e tasse dal quale emerge la posizione regolare dell’operatore economico (ai sensi degli artt. 94, comma 6, e 95, comma 2 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("accertamento della regolarità contributiva rispetto all’assolvimento degli obblighi relativi al versamento dei contributi previdenziali mediante acquisizione del Documento Unico di Regolarità Contributiva dal quale non risultano irregolarità relativamente al versamento dei contributi INPS e INAIL (ai sensi degli artt. 94, comma 6, e 95, comma 2 del Codice);")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("(eventuale) inserire le certificazioni/documenti acquisiti per la dimostrazione del possesso dei requisiti di cui all’art. 100 del D.Lgs. 36/2023 se richiesti;")), style = "Elenco punto")
    }
    doc <- doc |> 
      body_replace_text_at_bkm(bookmark = "bookmark_oe", paste0(Fornitore, " (P.IVA ", Fornitore..P.IVA, ", codice terzo SIGLA ", Fornitore..Codice.terzo.SIGLA, ")"))
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |> 
      cursor_reach("CAMPO.PROPONE") |>
      body_replace_all_text("CAMPO.PROPONE", ftext("la formalizzazione dell’affidamento diretto, immediatamente efficace, nei confronti del predetto operatore economico."), only_at_cursor = TRUE)
    }else{
      doc <- doc |> 
        cursor_reach("CAMPO.PROPONE") |>
        body_replace_all_text("CAMPO.PROPONE", "la formalizzazione dell’affidamento diretto nei confronti del predetto operatore economico.", only_at_cursor = TRUE)
    }
    doc <- doc |> 
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar("Il responsabile unico del progetto (RUP)", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "5 Atto istruttorio.docx"))
    
    cat("\014")
    cat("

    Documento '", pre.nome.file, "5 Atto istruttorio.docx' generato e salvato in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Importo.senza.IVA, RUP,Prot..RAS, Prot..provv..impegno, Pagina.web)
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
    
    ## Dich. Ass. RUP ----
    if(PNRR!="No"){
      download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.ras |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.rup), ftext(" "), ftext(dott.rup), ftext(" "), ftext(RUP, fpt.b), ftext(", "),
                         ftext(nato.rup), ftext(" "), ftext(RUP..Luogo.di.nascita), ftext(" il "),
                         ftext(RUP..Data.di.nascita), ftext(", codice fiscale "), ftext(RUP..Codice.fiscale), ftext(", ")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in relazione all'incarico di responsabile unico del progetto ("),
                         ftext("RUP", fpt.b),
                         ftext(") per l'affidamento "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext(") all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATE", fpt.b),
                         ftext(" le disposizioni di cui al decreto legislativo 8 aprile 2013 n. 39 in materia di incompatibilità e inconferibilità di incarichi presso le pubbliche amministrazioni e presso gli enti privati in controllo pubblico;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("consapevole delle responsabilità e delle sanzioni penali stabilite dalla legge per le false attestazioni e le dichiarazioni mendaci (artt. 75 e 76 D.P.R. n° 445/2000 e s.m.i.), sotto la propria responsabilità;")), style = "Normal") |>
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
      body_add_fpar(fpar("Il responsabile unico del progetto (RUP)", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2") |>
      body_add_break()
    print(doc, target = paste0(pre.nome.file, "4.4 Dichiarazione assenza conflitto RUP.docx"))
    
    ## Dich. Ass. RICH-TIT ----
    if(PNRR!="No"){
      download.file(paste(lnk, "Dich_conf_verso_tit.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.ras |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
    }
    doc <- doc |>
      cursor_reach("CAMPO.NOMINATIVO") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext(sottoscritto.ric), ftext(" "), ftext(dott.ric), ftext(" "), ftext(Richiedente, fpt.b), ftext(", "),
                         ftext(nato.ric), ftext(" "), ftext(Richiedente..Luogo.di.nascita), ftext(" il "),
                         ftext(Richiedente..Data.di.nascita), ftext(", codice fiscale "), ftext(Richiedente..Codice.fiscale),
                         ftext(", in qualità di "),
                         ftext("richiedente", fpt.b),
                         ftext(" l'affidamento diretto all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(", consapevole delle conseguenze penali di dichiarazioni mendaci, falsità in atti o uso di atti falsi, ai sensi dell’art. 76 D.P.R. 445/2000,")), style = "Normal") |>
      cursor_reach("CAMPO.DATA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar("Il richiedente l'affidamento", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.ric),
                         ftext(" "),
                         ftext(Richiedente),
                         ftext(")")), style = "Firma 2") |>
      cursor_reach("CAMPO.DATA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar("Il richiedente l'affidamento", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.ric),
                         ftext(" "),
                         ftext(Richiedente),
                         ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "4.6 Dichiarazione assenza conflitto tit_eff verso RICH.docx"))
    
    ## Dich. Ass. RESP-TIT ----
    if(Richiedente!=Responsabile.progetto){
      if(PNRR!="No"){
        download.file(paste(lnk, "Dich_conf_verso_tit.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- doc |>
          footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
        file.remove("tmp.docx")
        file.remove(logo)
      }else{
        doc <- doc.ras |>
          headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
        }
        file.remove("tmp.docx")
      }
      doc <- doc |>
        cursor_reach("CAMPO.NOMINATIVO") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext(sottoscritto.resp), ftext(" "), ftext(dott.resp), ftext(" "), ftext(Responsabile.progetto, fpt.b), ftext(", "),
                           ftext(nato.resp), ftext(" "), ftext(Responsabile.progetto..Luogo.di.nascita), ftext(" il "),
                           ftext(Responsabile.progetto..Data.di.nascita), ftext(", codice fiscale "), ftext(Responsabile.progetto..Codice.fiscale),
                           ftext(", in relazione alla procedura di affidamento diretto all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), "),
                           ftext("titolare dei fondi e responsabile", fpt.b),
                           ftext(" del "),
                           ftext(Progetto.int),
                           ftext(", consapevole delle conseguenze penali di dichiarazioni mendaci, falsità in atti o uso di atti falsi, ai sensi dell’art. 76 D.P.R. 445/2000,")), style = "Normal") |>
        cursor_reach("CAMPO.DATA") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar("Il responsabile unico del progetto (RUP)", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(Dott.rup),
                           ftext(" "),
                           ftext(RUP),
                           ftext(")")), style = "Firma 2") |>
        cursor_reach("CAMPO.DATA") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar("Il titolare dei fondi e responsabile del progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.resp),
                           ftext(" "),
                           ftext(Responsabile.progetto),
                           ftext(")")), style = "Firma 2")
      print(doc, target = paste0(pre.nome.file, "4.7 Dichiarazione assenza conflitto tit_eff verso RESP.docx"))
    }
    
    cat("

    Dichiarazioni di assenza conflitto di interesse generate e salvate in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Fornitore, Pagina.web)
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
    
    ## Dich. Ass. RSS-TIT ----
    if(PNRR!="No"){
      download.file(paste(lnk, "Dich_conf_verso_tit.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.ras |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
    }
    doc <- doc |>
      cursor_reach("CAMPO.NOMINATIVO") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext(sottoscritto.rss), ftext(" "), ftext(RSS, fpt.b), ftext(","),
                         ftext(nato.rss),
                         ftext(" "),
                         ftext(RSS.dich, fpt.b),
                         ftext(", in relazione all'affidamento diretto all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(", consapevole delle conseguenze penali di dichiarazioni mendaci, falsità in atti o uso di atti falsi, ai sensi dell’art. 76 D.P.R. 445/2000,")), style = "Normal") |>
      cursor_reach("CAMPO.DATA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(firma.RSS, run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(RSS),
                         ftext(")")), style = "Firma 2") |>
      cursor_reach("CAMPO.DATA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(firma.RSS, run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(RSS),
                         ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "4.8 Dichiarazione assenza conflitto tit_eff verso RSS.docx"))
    
    ## Dich. Ass. RUP-TIT ----
    if(PNRR!="No"){
      download.file(paste(lnk, "Dich_conf_verso_tit.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.ras |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
    }
    doc <- doc |>
      cursor_reach("CAMPO.NOMINATIVO") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext(sottoscritto.rup), ftext(" "), ftext(dott.rup), ftext(" "), ftext(RUP, fpt.b), ftext(", "),
                         ftext(nato.rup), ftext(" "), ftext(RUP..Luogo.di.nascita), ftext(" il "),
                         ftext(RUP..Data.di.nascita), ftext(", codice fiscale "), ftext(RUP..Codice.fiscale),
                         ftext(", in relazione all'incarico di responsabile unico del progetto ("),
                         ftext("RUP", fpt.b),
                         ftext(") per l'affidamento diretto all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(", consapevole delle conseguenze penali di dichiarazioni mendaci, falsità in atti o uso di atti falsi, ai sensi dell’art. 76 D.P.R. 445/2000,")), style = "Normal") |>
      cursor_reach("CAMPO.DATA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar("Il responsabile unico del progetto (RUP)", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2") |>
      cursor_reach("CAMPO.DATA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar("Il responsabile unico del progetto (RUP)", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "4.9 Dichiarazione assenza conflitto tit_eff verso RUP.docx"))
    
    ## Dich. Ass. SUP-TIT ----
    if(Supporto.RUP!=trattini){
      if(PNRR!="No"){
        download.file(paste(lnk, "Dich_conf_verso_tit.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- doc |>
          footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
        file.remove("tmp.docx")
        file.remove(logo)
      }else{
        doc <- doc.ras |>
          headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
        }
        file.remove("tmp.docx")
      }
      doc <- doc |>
        cursor_reach("CAMPO.NOMINATIVO") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext(sottoscritto.sup), ftext(" "), ftext(dott.sup), ftext(" "), ftext(Supporto.RUP, fpt.b), ftext(", "), 
                           ftext(nato.sup), ftext(" "), ftext(Supporto.RUP..Luogo.di.nascita),
                           ftext(" il "), ftext(Supporto.RUP..Data.di.nascita),
                           ftext(", codice fiscale "), ftext(Supporto.RUP..Codice.fiscale),
                           ftext(", in relazione all'incarico di "),
                           ftext("supporto al RUP", fpt.b),
                           ftext(" per l'affidamento diretto all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), nell'ambito del "),
                           ftext(Progetto.int),
                           ftext(", consapevole delle conseguenze penali di dichiarazioni mendaci, falsità in atti o uso di atti falsi, ai sensi dell’art. 76 D.P.R. 445/2000,")), style = "Normal") |>
        cursor_reach("CAMPO.DATA") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar("Il supporto al RUP", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.sup),
                           ftext(" "),
                           ftext(Supporto.RUP),
                           ftext(")")), style = "Firma 2") |>
        cursor_reach("CAMPO.DATA") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar("Il supporto al RUP", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.sup),
                           ftext(" "),
                           ftext(Supporto.RUP),
                           ftext(")")), style = "Firma 2")
      print(doc, target = paste0(pre.nome.file, "4.10 Dichiarazione assenza conflitto tit_eff verso SUP.docx"))
    }
    
    cat("

    Dichiarazioni di assenza conflitto di interesse generate e salvate in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Fornitore, Pagina.web)
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
  
  # DaC PNRR ----
  dac.pnrr <- function(){
    cat("\014")
    
    if(PNRR!="No"){
      download.file(paste(lnk, "DaC.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.dac |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
      file.remove("tmp.docx")
      file.remove(logo)
    }
    
    doc <- doc |>
      cursor_reach("CAMPO.DELLA.FORNITURA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("OGGETTO", fpt.b), ftext(": DECISIONE DI CONTRATTARE PER L’AFFIDAMENTO DIRETTO "),
                         ftext(toupper(della.fornitura)), ftext(" DI “"),
                         ftext(toupper(Prodotto), fpt.b),
                         ftext("”, ORDINE "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b), ftext(y, fpt.b),
                         ftext(", CIG "),
                         ftext(CIG, fpt.b),
                         ftext(" ("),
                         ftext(toupper(Pagina.web)),
                         ftext("), NELL'AMBITO DEL "),
                         ftext(toupper(Progetto.int)),
                         ftext(".")), style = "Normal") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      cursor_reach("CAMPO.NOMINE") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")
    
    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(avviso.pnrr)), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr)), style = "Normal")
    if(PNRR=="onFoods Spoke 4"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr2)), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTI", fpt.b), ftext(decreto.pnrr3)), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr4)), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(decreto.pnrr5)), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr6)), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(decreto.pnrr7)), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(decreto.pnrr8)), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(decreto.pnrr9)), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(decreto.pnrr10)), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(decreto.pnrr11)), style = "Normal") 
    }
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la "), ftext("richiesta di acquisto prot. ", fpt.b),
                         ftext(Prot..RAS, fpt.b), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”, nell’ambito delle attività previste dal "),
                         ftext(Progetto.cup),
                         ftext(", corredata dal preventivo d'importo pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA, formulato dall'operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), "),
                         ftext(preventivo.individuato)), style = "Normal") |>
      cursor_reach("CAMPO.NOMINA.RUP") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento prot. n. "),
                         ftext(Prot..provv..impegno),
                         ftext(", con il quale è stato nominato "),
                         ftext(il.dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(" quale Responsabile Unico del Progetto (RUP) ai sensi dell’art. 15 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato "),
                         ftext(paste0(Fornitore, " (P.IVA ", Fornitore..P.IVA, ")")),
                         ftext(" ha presentato, attraverso la piattaforma telematica di negoziazione, un’offerta ritenuta congrua, corredata dalle dichiarazioni sostitutive richieste, in merito al possesso dei requisiti prescritti dalla S.A., d’importo uguale o inferiore rispetto a quello stimato dalla stazione appaltante, pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" l’"), ftext("istruttoria del RUP prot. n. ", fpt.b),
                         ftext(Prot..atto.istruttorio, fpt.b),
                         ftext(", contenente l’esito positivo relativo alla verifica dei requisiti dichiarati in sede di affidamento dall’operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nonché la proposta di affidamento diretto al medesimo operatore economico "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO ", fpt.b),
                         ftext(rotazione.individuata)), style = "Normal") |>
      cursor_reach("CAMPO.DISPONIBILITA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b), ftext(" la disponibilità finanziaria per la copertura della spesa sui fondi del già richiamato progetto, allocati al GAE "),
                         ftext(GAE),
                         ftext(", voce del piano "), ftext(Voce.di.spesa), ftext(";")), style = "Normal")
    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        cursor_bookmark("bookmark_procedere") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext("DI PROCEDRE", fpt.b),
                           ftext(" all'acquisizione "),
                           ftext(della.fornitura),
                           ftext(" di cui trattasi, mediante affidamento diretto all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext(", codice terzo SIGLA "),
                           ftext(Fornitore..Codice.terzo.SIGLA),
                           ftext("), con sede legale in "),
                           ftext(Fornitore..Sede),
                           ftext(", che ha presentato il proprio preventivo ammontante a "),
                           ftext(Importo.senza.IVA, fpt.b),
                           ftext(" oltre IVA;")), style = "Elenco numero")
        # body_replace_text_at_bkm(bookmark = "bookmark_procedere", paste0("all'acquisizione ", della.fornitura, " di cui trattasi, mediante affidamento diretto all'operatore economico ", all.OE, 
        #                                             ", P.IVA ", Fornitore..P.IVA, ", codice terzo SIGLA ", Fornitore..Codice.terzo.SIGLA, ", con sede legale in ",
        #                                             Fornitore..Sede,
        #                                             ", che ha presentato il proprio preventivo, ammontante a ",
        #                                             Importo.senza.IVA,
        #                                             " oltre IVA;"))
    }else{
      doc <- doc |>
        cursor_bookmark("bookmark_procedere") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext("DI PROCEDRE", fpt.b),
                           ftext(" all'acquisizione "),
                           ftext(della.fornitura),
                           ftext(" di cui trattasi, mediante affidamento diretto, immediatamente efficace, all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext(", codice terzo SIGLA "),
                           ftext(Fornitore..Codice.terzo.SIGLA),
                           ftext("), con sede legale in "),
                           ftext(Fornitore..Sede),
                           ftext(", che ha presentato il proprio preventivo ammontante a "),
                           ftext(Importo.senza.IVA, fpt.b),
                           ftext(" oltre IVA;")), style = "Elenco numero")
    }
    doc <- doc |>
      cursor_reach("CAMPO.DI.IMPEGNARE") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("impegno di spesa n. "),
                         ftext(N..impegno.di.spesa),
                         ftext(" di "),
                         ftext(Importo.con.IVA),
                         ftext(" IVA inclusa sui fondi del progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce del piano "),
                         ftext(Voce.di.spesa),
                         ftext(", sul GAE "),
                         ftext(GAE),
                         ftext(";")), style = "Elenco punto liv2")
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("€ 35,00, Voce del piano 13096 “Pubblicazione bandi di gara” sul GAE [completare] per la quota stazione appaltante della contribuzione ANAC;")), style = "Elenco punto liv2")
    } 
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        cursor_reach("DI STABILIRE altresì che, trattandosi di affidamento d’importo inferiore a") |>
        body_remove()
    }
    doc <- doc |>
      cursor_reach("CAMPO.FIRMA") |>
      body_remove() |>
      body_add_par("Visto di regolarità contabile.", style = "Firma 1") |>
      #body_add_par(resp.segr, style = "Firma 1") |>
      #body_add_fpar(fpar(ftext("("), ftext(RAMM), ftext(")")), style = "Firma 1") |>
      body_add_par("La segretaria amministrativa", style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(sig.ra Concetta Mottura)")), style = "Firma 1") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "6 Decisione a contrattare.docx"))
    
    cat("\014")
    cat("

    Documento '", pre.nome.file, "6 Decisione a contrattare.docx' generato e salvato in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Importo.senza.IVA, Voce.di.spesa, GAE, RUP, Prot..atto.istruttorio, Pagina.web)
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
  
  # Lettera d'ordine PNRR ----
  ldo.pnrr <- function(){
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
    
    download.file(paste(lnk, "LdO.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    if(PNRR!="No"){
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
        headers_replace_text_at_bkm(bookmark = "bookmark_headers", toupper(Progetto.int))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers", sede1)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("LETTERA D’ORDINE "), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
      body_replace_text_at_bkm(bookmark = "bookmark_cup_it", CUP2) |>
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
      body_replace_text_at_bkm("bookmark_fatturazione2", dicitura.fatturazione) |>
      cursor_reach("CAMPO.FIRMA.LDO.IT")
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Il contraente, a garanzia dell’adempimento di tutte le obbligazioni del presente atto, ha costituito la garanzia definitiva ____________________ di € __________ (euro __________), mediante cauzione/fideiussione agli atti.")), style = "Normal", pos = "on")
    }
    doc <- doc |>
      body_add_fpar(fpar(""), style = "Normal", pos = "on") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    
    if(PNRR!='No'){
      doc <- doc |>
        cursor_reach("CAMPO.FATTURAZIONE.2") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext("Fatturazione", fpt.b), ftext(": la fattura, redatta secondo la normativa vigente, dovrà riportare, pena il rifiuto della stessa, il numero d'ordine, il numero di protocollo (si veda in alto nella pagina della lettera d'ordine), il CIG, il CUP e la seguente dicitura: '"),
                           ftext(dicitura.fattura, fpt.b),
                           ftext("'.")), style = "Elenco punto")
    }else{
      doc <- doc |>
        cursor_reach("CAMPO.FATTURAZIONE.2") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext("Fatturazione", fpt.b), ftext(": la fattura, redatta secondo la normativa vigente, dovrà riportare, pena il rifiuto della stessa, il numero d'ordine, il numero di protocollo (si veda in alto nella pagina della lettera d'ordine), il CIG e il CUP.")), style = "Elenco punto")
    }
    
    if(PNRR!='No'){
      if(Tipo.acquisizione=='Beni'){
        doc <- doc |>
          cursor_bookmark("bookmark_conformita") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Verifica di conformità", fpt.b), ftext(": la presente fornitura è soggetta a verifica di conformità da effettuarsi, secondo quanto previsto dall’art. 116 e nell’Allegato II.14 del codice dei contratti entro 2 mesi. A seguito della verifica di conformità si procede al pagamento della rata di saldo e, se prevista, allo svincolo della cauzione.")), style = "Elenco punto")
      }else{
        doc <- doc |>
          cursor_bookmark("bookmark_conformita") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Verifica di regolare esecuzione", fpt.b), ftext(": La stazione appaltante, per il tramite del RUP, emette il certificato di regolare esecuzione, secondo le modalità indicate nell'Allegato II.14 al codice dei contratti pubblici, entro _____ mesi. A seguito dell’emissione del certificato di regolare esecuzione si procede al pagamento della rata di saldo e, se prevista, allo svincolo della cauzione.")), style = "Elenco punto")
      } 
    }else{
      doc <- doc |>
        cursor_bookmark("bookmark_conformita") |>
        body_remove() |>
        cursor_backward()
    }
    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Clausola risolutiva espressa", fpt.b), ftext(": l’ordine è emesso in applicazione delle disposizioni contenute nell’art. 52, commi 1 e 2 del d.lgs 36/2023. Il CNR ha diritto di risolvere il contratto/ordine in caso di accertamento della carenza dei requisiti di partecipazione. Per la risoluzione del contratto trovano applicazione l’art. 122 del d.lgs. 36/2023, nonché gli articoli 1453 e ss. del Codice Civile. Il CNR darà formale comunicazione della risoluzione al fornitore, con divieto di procedere al pagamento dei corrispettivi, se non nei limiti delle prestazioni già eseguite.")), style = "Elenco punto")
    }
    
    if(Fornitore..Nazione=="Italiana"){
      b <- cursor_reach(doc, "CAMPO.INIZIO.LDO.EN")
      b <- doc$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which -5
      doc <- cursor_reach(doc, "CAMPO.FIRMA.LDO.EN")
      for(i in 1:(e-b)){
        doc <- body_remove(doc)
      }
    }else{
      doc <- doc |>
        cursor_reach("CAMPO.INIZIO.LDO.EN") |>
        body_add_fpar(fpar(ftext("PURCHASE ORDER "), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
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
        body_replace_text_at_bkm("bookmark_fatturazione2_eng", dicitura.fatturazione.eng) |>
        cursor_reach("CAMPO.FIRMA.LDO.EN")
      if(Importo.senza.IVA.num>=40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("The contractor, as a guarantee of the fulfillment of all the obligations of this deed, has constituted the definitive guarantee ____________________ of € __________ (euro __________), by means of a security/guarantee of the documents.")), style = "Normal", pos = "on")
      }
      doc <- doc |>
        body_add_fpar(fpar(""), style = "Normal", pos = "on") |>
        body_add_fpar(fpar(ftext("The Responsible")), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
      
      if(PNRR!="No"){
        doc <- doc |>
          cursor_reach("CAMPO.FATTURAZIONE.2") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Invoice", fpt.b), ftext(": the invoice, drawn up in accordance with current legislation, must include, under penalty of rejection, the purchase order number, the number of the registration protocol (see on the top of the purchase order page), the CIG, the CUP and the following phrase: '"),
                             ftext(dicitura.fattura, fpt.b),
                             ftext("'.")), style = "Elenco punto 2")
      }else{
        doc <- doc |>
          cursor_reach("CAMPO.FATTURAZIONE.2") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Invoice", fpt.b), ftext(": the invoice, drawn up in accordance with current legislation, must include, under penalty of rejection, the purchase order number, the number of the registration protocol (see on the top of the purchase order page), the CIG and the CUP.")), style = "Elenco punto 2")
      }
      
      if(PNRR!='No'){
        if(Tipo.acquisizione=='Beni'){
          doc <- doc |>
            cursor_reach("CAMPO.VERIFICA.CONFORMITA") |>
            body_remove() |>
            cursor_backward() |>
            body_add_fpar(fpar(ftext("Verification of conformity", fpt.b), ftext(": this supply is subject to verification of conformity to be carried out, in accordance with the provisions of art. 116 and in Annex II.14 of the Public Contracts Code within 2 months. Following the verification of conformity, the balance instalment is paid and, if constituted, the security deposit is released.")), style = "Elenco punto 2")
        }else{
          doc <- doc |>
            cursor_reach("CAMPO.VERIFICA.CONFORMITA") |>
            body_remove() |>
            cursor_backward() |>
            body_add_fpar(fpar(ftext("Verification of regular execution", fpt.b), ftext(": the contracting authority, through the RUP, issues the certificate of regular execution, in accordance with the methods indicated in Annex II.14 of the Public Contracts Code, within ___ months. Following the issuance of the certificate of regular execution, the balance instalment is paid and, if constituted, the security deposit is released.")), style = "Elenco punto 2")
        }
      }else{
        doc <- doc |>
          cursor_reach("CAMPO.VERIFICA.CONFORMITA") |>
          body_remove() |>
          cursor_backward()
      }
      if(Importo.senza.IVA.num<40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Express termination clause", fpt.b), ftext(": the order is issued in application of the provisions contained in art. 52, paragraphs 1 and 2 of Legislative Decree 36/2023. The CNR has the right to terminate the contract/order in the event of a lack of participation requirements being ascertained. For the termination of the contract, art. 122 of Legislative Decree 36/2023, as well as articles 1453 et seq. of the Civil Code, apply. The CNR will formally communicate the termination to the supplier, with a ban on proceeding with the payment of the fees, except within the limits of the services already performed.")), style = "Elenco punto 2")
      }
    }
    print(doc, target = paste0(pre.nome.file, "7 Lettera ordine.docx"))
    
    cat("\014")
    cat("

    Documento '", pre.nome.file, "7 Lettera ordine.docx' generato e salvato in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, CIG, RUP, RDO, Fornitore, Fornitore..Sede, Fornitore..P.IVA, Fornitore..PEC, Fornitore..E.mail, Importo.senza.IVA, Aliquota.IVA, Richiedente..Luogo.di.consegna, Pagina.web, Prot..DaC)
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
  
  # Prestazione resa PNRR ----
  dic_pres.pnrr <- function(){
    if(PNRR!="No"){
      download.file(paste(lnk, "Vuoto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      doc <- doc.dic.pres |>
        headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      }
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
    
    Imponibile.ldo <- colnames(pr)[7]
    IVA.ldo <- pr[1,7]
    Importo.ldo <- pr[2,7]
    Imponibile.ldo.txt <- paste("€", format(as.numeric(Imponibile.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    IVA.ldo.txt <- paste("€", format(as.numeric(IVA.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    Importo.ldo.txt <- paste("€", format(as.numeric(Importo.ldo), format='f', digits=2, nsmall=2, big.mark = ".", decimal.mark = ","))
    
    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_par("DICHIARAZIONE DI PRESTAZIONE RESA", style = "heading 1", pos = "on") |>
      body_add_par("Il responsabile unico del progetto (RUP)", style = "heading 1") |>
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
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(";")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTA ", fpt.b),
                         ftext(ordine.trattativa.scelta.pres)), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la lettera d’ordine "), ftext(sede),
                         ftext(" "), ftext(ordine), ftext(y),
                         ftext(" di "), ftext(Importo.ldo.txt),
                         ftext(" IVA inclusa emessa nei confronti dell'operatore economico "),
                         #ftext(Prot..lettera.ordine),
                         ftext(Fornitore), ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext("; codice terzo SIGLA "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(");")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il documento di trasporto;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di aver svolto la procedura secondo la normativa vigente;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext(materiale.conforme)), style = "Elenco punto") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Il responsabile unico del progetto (RUP)")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(dott.rup), ftext(" "), ftext(RUP), ftext(")")), style = "Firma 2")
      
    print(doc, target = paste0(pre.nome.file, "8 Dichiarazione prestazione resa.docx"))

    cat("

    Documento '", pre.nome.file, "8 Prestazione resa.docx' generato e salvato in ", pat)
    
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
  
  # Doppio finanziamento ----
  doppio_fin.pnrr <- function(){
    download.file(paste(lnk, "Vuoto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
    file.remove("tmp.docx")
    file.remove(logo)
    
    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_par("DICHIARAZIONE SOSTITUTIVA DI ASSENZA DOPPIO FINANZIAMENTO", style = "heading 1", pos = "on") |>
      body_add_par("ai sensi degli artt. 46 e 47 del D.P.R. 28 dicembre 2000, n. 445", style = "heading 1") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Il sottoscritto dott. Francesco Di Serio, nato a Cava de’ Tirreni (SA) il 29/09/1965, codice fiscale DSRFNC65P29C361R, direttore dell'IPSP a decorrere dal giorno 1/5/2024 per quattro anni in base al provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, in relazione all'affidamento diretto "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(ordine),
                         ftext(y),
                         ftext(", "),
                         ftext("CIG ", fpt.b),
                         ftext(CIG, fpt.b),
                         ftext(" ("),
                         ftext(Pagina.web),
                         ftext("), decisione a contrattare prot. n. "),
                         ftext(Prot..DaC),
                         ftext(", all'operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del "),
                         ftext(Progetto.int),
                         ftext(", consapevole della responsabilità penale cui può andare incontro in caso di dichiarazione falsa o comunque non corrispondente al vero (art. 76 del D.P.R. n. 445 del 28/12/2000), ai sensi del D.P.R. n. 445 del 28/12/2000 e ss.mm.ii.")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTA ", fpt.b),
      #                    ftext(ordine.trattativa.scelta.pres)), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("l’"),
      ftext("assenza del doppio finanziamento", fpt.b),
      ftext(" a valere su fonti di finanziamento pubbliche anche di diversa natura, come specificato dalla Circolare n. 33 del 31 dicembre 2021 del Ministero dell’Economia e delle Finanze.")), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Il direttore dell'IPSP")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("(dott. Francesco Di Serio)")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "9 Dichiarazione assenza doppio finanziamento.docx"))
    
    #cat("\014")
    cat("

    Documento '", pre.nome.file, "9 Dichiarazione assenza doppio finanziamento.docx' generato e salvato in ", pat)
  }
  
  # Funzionalità bene ----
  fun_bene.pnrr <- function(){
    if(Inventariabile=='Inventariabile'){
      download.file(paste(lnk, "Vuoto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
      
      doc <- doc |>
        cursor_begin() |>
        cursor_forward() |>
        body_add_par("DICHIARAZIONE DI FUNZIONALITÀ DEL BENE", style = "heading 1", pos = "on") |>
        body_add_par("ai sensi degli artt. 46 e 47 del D.P.R. 28 dicembre 2000, n. 445", style = "heading 1") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext("Il sottoscritto "), ftext(dott.ric), ftext(" "), ftext(Richiedente), ftext(", "),
                           ftext(nato.ric), ftext(" "), ftext(Richiedente..Luogo.di.nascita), ftext(" il "),
                           ftext(Richiedente..Data.di.nascita), ftext(", codice fiscale "), ftext(Richiedente..Codice.fiscale), ftext(", in merito allo strumento “"),
                           ftext(Prodotto, fpt.b),
                           ftext("”, acquisito con affidamento diretto, ordine "),
                           ftext(ordine, fpt.b),
                           ftext(y, fpt.b),
                           ftext(", "),
                           ftext("CIG ", fpt.b),
                           ftext(CIG, fpt.b),
                           ftext(" ("),
                           ftext(Pagina.web),
                           ftext("), decisione a contrattare prot. n. "),
                           ftext(Prot..DaC),
                           ftext(", dall'operatore economico "),
                           ftext(Fornitore),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), nell'ambito del "),
                           ftext(Progetto.int),
                           ftext(", ")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("che le apparecchiature sono di fondamentale importanza per lo svolgimento delle attività del progetto in relazione allo scopo di mantenere aggiornate ed efficienti le apparecchiature scientifiche indispensabili per lo svolgimento delle azioni di ricerca programmate;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("che l’acquisto è perfettamente allineato con le previsioni di spesa del progetto e con la relativa ripartizione delle disponibilità economiche come previsto in fase di costruzione e redazione del progetto stesso;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("che, tutto questo considerato, le apparecchiature sono funzionali alle attività del progetto e verranno utilizzate esclusivamente per il conseguimento degli obiettivi realizzativi dello stesso.")), style = "Elenco punto") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext("Il direttore dell'IPSP")), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(dott. Francesco Di Serio)")), style = "Firma 2")
      
      print(doc, target = paste0(pre.nome.file, "10 Dichiarazione funzionalità bene.docx"))
      
      cat("

    Documento '", pre.nome.file, "10 Dichiarazione funzionalità bene.docx' generato e salvato in ", pat)
    }
  }
 
  # Checklist ----
  chklst.pnrr <- function(){
    download.file(paste(lnk, "Checklist.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
    file.remove("tmp.docx")
    file.remove(logo)
    
    doc <- doc |>
      body_replace_text_at_bkm("investimento", investimento) |>
      body_replace_text_at_bkm("investimento2", sub("(.*[0-9]\\.[0-9]) .*", "\\1", investimento)) |>
      body_replace_text_at_bkm("intervento", intervento) |>
      body_replace_text_at_bkm("attuatore", attuatore) |>
      body_replace_text_at_bkm("avvio", avvio) |>
      body_replace_text_at_bkm("costo", costo.totale) |>
      body_replace_text_at_bkm("costo_ammesso", costo.ammesso) |>
      body_replace_text_at_bkm("sede", sede2) |>
      body_replace_text_at_bkm("sede2", sede2) |>
      body_replace_all_text("CAMPO.DAC", Prot..DaC, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.IMPORTO", Importo.senza.IVA, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.IVA", IVA, only_at_cursor = FALSE) |>
      body_replace_text_at_bkm("prodotto", Prodotto) |>
      body_replace_text_at_bkm("fornitore", paste0(Fornitore, " (P.IVA ", Fornitore..P.IVA, ")")) |>
      #cursor_begin() |>
      body_replace_all_text("CAMPO.ISTRUTTORIA", Prot..atto.istruttorio, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.LDO", Prot..lettera.ordine, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.RUP", Prot..provv..impegno, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.DOCOE", Prot..atto.istruttorio, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.DOPPIOFIN", Prot..doppio.finanziamento, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.CIG", CIG, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.CUP", CUP, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.RDO", as.character(RDO), only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.PAGINA", Pagina.web, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.DECRETO", decreto.concessione, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.CODICE.PROGETTO", codice.progetto, only_at_cursor = FALSE)
    if(Tipo.acquisizione!="Beni"){
      doc <- doc |>
        body_replace_text_at_bkm("durata_affidamento", durata.affidamento) |>
        body_replace_text_at_bkm("durata_affidamento2", durata.affidamento)
    }
    
    print(doc, target = paste0(pre.nome.file, "11 Checklist.docx"))

    cat("

    Documento '", pre.nome.file, "11 Checklist.docx' generato e salvato in ", pat)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prot..DaC, Prot..atto.istruttorio, Prot..lettera.ordine, Prot..provv..impegno, Pagina.web)
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
  
    # Input ----
  answ <- function(){
    cat("\014")
    cat("

      ***************************
      *** BENVENUTI in AppOst ***
      ***************************

    Che documento vuoi generare?
      1: RAS, con eventuale avviso pubblico, Richiesta pagina web
      2: Provvedimento d'impegno, Decisione a contrattare
      3: Comunicazione CIG, Autocertificazioni operatore economico, Atto istruttorio, Lettera d'ordine, Prestazione resa, Provvedimento di liquidazione
      
      
    Solo per PNRR e PRIN:
      4: RAS, Assenza conflitto interesse, Richiesta pagina web
      5: Nomina RUP, Assenza conflitto interesse, Autocertificazioni operatore economico
      6: Atto istruttorio, Assenza conflitto interesse, Comunicazione CIG
      7: Decisione a contrattare, Assenza doppio finanziamento, Funzionalità del bene
      8: Lettera d'ordine, Prestazione resa
      9: Provvedimento di liquidazione, Checklist

")
      
    inpt <- readline()
    if(inpt==1){ras();pag()}
    if(inpt==2){provv_imp();dac()}
    if(inpt==3){com_cig();docoe();ai();ldo();dic_pres();provv_liq()}
    if(inpt==4){ras.pnrr();pag()}
    if(inpt==5){rup.pnrr();docoe.pnrr()}
    if(inpt==6){ai.pnrr();com_cig()}
    if(inpt==7){dac.pnrr();doppio_fin.pnrr();fun_bene.pnrr()}
    if(inpt==8){ldo.pnrr();dic_pres.pnrr()}
    if(inpt==9){provv_liq();chklst.pnrr()}
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
