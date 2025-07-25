appost <- function(){
  # pat <- utils::choose.dir(caption = "Seleziona la cartella dell'ordine")
  # setwd(pat)

  # Carica dati ordine ----
  #ultimo aggiornamento: ", format(Sys.Date(), "%d %B %Y"), "
  cat("\014")
  cat(paste0("

      ***************************
      *** BENVENUTI in AppOst ***
      ***************************
      
      AppOst è ottimizzata per affidamenti diretti di forniture e servizi <40.000 € ...
      e ora anche per:
      - servizi di natura non intellettuale;
      - forniture con posa in opera.
      
      Invece, per i seguenti ordini sono necessari adattamenti:
      - ordini >40.000 €
      - procedure aperte o negoziate

Digitare il numero d'ordine e premere INVIO caricare il file 'Ordini.csv' scaricato da Teams


      "))
    # oppure digitare '0' (zero) per scaricare il file 'Elenco prodotti.xlsx'
  # (da compilare prima di generare RAS e lettera d'ordine)
  #ordine <- "20_2024"
  #ordine <- 28
  ordine <- readline()

  if(ordine==0){
    # pat <- utils::choose.dir()
    # setwd(pat)
    download.file("https://raw.githubusercontent.com/giovabubi/appost/main/models/Elenco%20prodotti.xlsx", destfile = "Elenco prodotti.xlsx", method = "curl")
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
  #colnames(ordini)[3] <- "Data"
  ordini$Fornitore..P.IVA <- as.character(ordini$Fornitore..P.IVA)
  ordini$CPV <- NULL
  ordini$CPV <- as.character(ordini$CPV..CPV)
  ordini$Importo.senza.IVA.num <- sub(",(..)$", "_\\1", ordini$Importo.senza.IVA)
  ordini$Importo.senza.IVA.num <- gsub("\\.", "", ordini$Importo.senza.IVA.num)
  ordini$Importo.senza.IVA.num <- gsub("_", ".", ordini$Importo.senza.IVA.num)
  ordini$Importo.senza.IVA.num <- as.numeric(ordini$Importo.senza.IVA.num)
  ordini$Manodopera <- ifelse(is.na(ordini$Manodopera), "0,00 €", ordini$Manodopera)
  ordini$Manodopera.num <- sub(",(..)..$", "_\\1", ordini$Manodopera)
  ordini$Manodopera.num <- gsub("\\.", "", ordini$Manodopera.num)
  ordini$Manodopera.num <- gsub("_", ".", ordini$Manodopera.num)
  ordini$Manodopera.num <- as.numeric(ordini$Manodopera.num)
  ordini$Oneri.sicurezza <- ifelse(is.na(ordini$Oneri.sicurezza), "0,00 €", ordini$Oneri.sicurezza)
  ordini$Oneri.sicurezza.num <- sub(",(..)..$", "_\\1", ordini$Oneri.sicurezza)
  ordini$Oneri.sicurezza.num <- gsub("\\.", "", ordini$Oneri.sicurezza.num)
  ordini$Oneri.sicurezza.num <- gsub("_", ".", ordini$Oneri.sicurezza.num)
  ordini$Oneri.sicurezza.num <- as.numeric(ordini$Oneri.sicurezza.num)

  sc <- subset(ordini, ordini$Ordine.N.==ordine)
  
  anno24 <- grep("_2024", sc$Ordine.N.)
  anno25 <- grep("_2025", sc$Ordine.N.)
  ordine.orig <- ordine
  if(length(anno24==1)){
    y <- "/2024"
    y2 <- 2024
    sc$Ordine.N. <- sub("_2024", "", sc$Ordine.N.)
    ordine <- sub("_2024", "", ordine)
  }else if(length(anno25==1)){
    y <- "/2025"
    y2 <- 2025
    sc$Ordine.N. <- sub("_2025", "", sc$Ordine.N.)
    ordine <- sub("_2025", "", ordine)
  }else{
    y <- "/2025"
    y2 <- 2025
  }
  
  sc$Aliquota.IVA.num <- as.numeric(ifelse(sc$Aliquota.IVA=='22%', 0.22,
                                           ifelse(sc$Aliquota.IVA=='10%', 0.1,
                                                  ifelse(sc$Aliquota.IVA=='4%', 0.04, 0))))
  sc$IVA <- sc$Importo.senza.IVA.num * sc$Aliquota.IVA.num
  sc$IVA.num <- sc$IVA
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
    RSS.nome <- "Giovanni Nicola Bubici"
    RSS <- paste("dott.", RSS.nome)
    RSS.email <- 'giovanninicola.bubici@cnr.it'
    RAMM <- 'Dott. Nicola Centorame'
    RAMM.email <- 'nicola.centorame@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Bari"
    firma.RSS <- "Il responsabile della sede secondaria di Bari dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, via G. Amendola 122/D, 70126 Bari, Italia."
    nomina.RSS <- "3903 dell'8/1/2025 di nomina del dott. Giovanni Nicola Bubici quale Responsabile della Sede Secondaria di Bari dell’IPSP per il periodo dall’8/1/2025 al 31/12/2025"
    nomina.RSS2 <- "3903 dell'8/1/2025 relativo al conferimento dell’incarico del dott. Giovanni Nicola Bubici a Responsabile delegato alla gestione della sede secondaria di Bari dell’IPSP con decorrenza dall'8/1/2025 al 31/12/2025;"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Nicola Centorame quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "Il responsabile amministrativo della sede secondaria di Bari dell'IPSP"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a Foggia il giorno 11/11/1977, codice fiscale BBCGNN77S11D643H,"
    RSS.dich <- "responsabile della sede secondaria di Bari dell'IPSP"
    CUU <- "4S488Q"
    cdr <- paste("CdR 121.001.000 IPSP", sede2)
  }else if(sede=='TO'){
    sede1 <- 'Torino'
    sede2 <- 'Sede Secondaria di Torino'
    RSS.nome <- 'Stefano Ghignone'
    RSS <- paste("dott.", RSS.nome)
    RSS.email <- 'stefano.ghignone@cnr.it'
    RAMM <- "Dott.ssa Lucia Allione"
    RAMM.email <- 'lucia.allione@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Torino"
    firma.RSS <- "Il responsabile della sede secondaria di Torino dell'IPSP"
    in.qualita.RSS <- "responsabile della sede secondaria di Torino dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, viale Mattioli, 25, 10125 Torino, Italia."
    nomina.RSS <- "3906 dell'8/1/2025 di nomina del dott. Stefano Ghignone quale Responsabile della Sede Secondaria di Torino dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RSS2 <- "3903 dell'8/1/2025 relativo al conferimento dell’incarico del dott. Stefano Ghignone a Responsabile delegato alla gestione della sede secondaria di Torino dell’IPSP con decorrenza dall'8/1/2025 al 31/12/2025;"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della dott.ssa Lucia Allione quale Responsabile Amministrativo della Sede Secondaria di Torino dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La responsabile amministrativa della sede secondaria di Torino dell'IPSP"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a Chieri (TO) il 29/5/1972, codice fiscale GHGSFN72E29C627M,"
    RSS.dich <- "responsabile della sede secondaria di Torino dell'IPSP"
    CUU <- "PE2U6Q"
    cdr <- paste("CdR 121.00_.000 IPSP", sede2)
  }else if(sede=='NA'){
    sede1 <- 'Portici'
    sede2 <- 'Sede Secondaria di Portici'
    RSS.nome <- 'Michelina Ruocco'
    RSS <- paste("dott.ssa", RSS.nome)
    RSS.email <- 'michelina.ruocco@cnr.it'
    RAMM <- 'Dott. Ettore Magaldi'
    RAMM.email <- 'ettore.magaldi@ipsp.cnr.it'
    al.RSS <- "Alla responsabile della sede secondaria di Portici"
    firma.RSS <- "La responsabile della sede secondaria di Portici dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, piazzale Enrico Fermi, 1, 80055 Portici (NA), Italia."
    nomina.RSS <- "3907 dell'8/1/2025 di nomina della dott.ssa Michelina Ruocco quale Responsabile della Sede Secondaria di Portici dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RSS2 <- "3903 dell'8/1/2025 relativo al conferimento dell’incarico della dott.ssa Michelina Ruocco a Responsabile delegata alla gestione della sede secondaria di Portici dell’IPSP con decorrenza dall'8/1/2025 al 31/12/2025;"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina del dott. Ettore Magaldi quale Responsabile Amministrativo della Sede Secondaria di Bari dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "Il responsabile amministrativo"
    sottoscritto.rss <- "La sottoscritta "
    nato.rss <- " nata a Sant'Agnello il 28/05/1969, codice fiscale RCCMHL69E68I208P,"
    RSS.dich <- "responsabile della sede secondaria di Portici dell'IPSP"
    cdr <- paste("CdR 121.00_.000 IPSP", sede2)
    CUU <- "YOBWQ2"
  }else if(sede=='FI'){
    sede1 <- 'Sesto Fiorentino'
    sede2 <- 'Sede Secondaria di Sesto Fiorentino'
    RSS.nome <- "Nicola Luchi"
    RSS <- paste("dott.", RSS.nome)
    RSS.email <- "nicola.luchi@ipsp.cnr.it"
    RAMM <- "Sig.ra Francesca Pesciolini"
    RAMM.email <- 'francesca.pesciolini@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Sesto Fiorentino"
    firma.RSS <- "Il responsabile della sede secondaria di Sesto Fiorentino dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, via Madonna del Piano, 10, 50019 Sesto F.no (FI), Italia."
    nomina.RSS <- "3904 dell'8/1/2025 di nomina del dott. Nicola Luchi quale Responsabile della Sede Secondaria di Sesto Fiorentino dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RSS2 <- "3903 dell'8/1/2025 relativo al conferimento dell’incarico del dott. Nicola Luchi a Responsabile delegato alla gestione della sede secondaria di Sesto Fiorentino dell’IPSP con decorrenza dall'8/1/2025 al 31/12/2025;"
    nomina.RAMM <- "146220 del 2/5/2024 di nomina della sig.ra Francesca Pesciolini quale Responsabile Amministrativo della Sede Secondaria di Sesto Fiorentino dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La responsabile amministrativa"
    sottoscritto.rss <- "Il sottoscritto "
    nato.rss <- " nato a Firenze il 27/12/1975, codice fiscale LCHNCL75T27D612B,"
    RSS.dich <- "responsabile della sede secondaria di Sesto Fiorentino dell'IPSP"
    CUU <- "KWH4FD"
    cdr <- paste("CdR 121.00_.000 IPSP", sede2)
  }else if(sede=='PD'){
    sede1 <- 'Legnaro'
    sede2 <- 'Sede Secondaria di Legnaro'
    RSS.nome <- "Laura Scarabel"
    RSS <- paste("dott.ssa", RSS.nome)
    RSS.email <- "laura.scarabel@ipsp.cnr.it"
    RAMM <- "Dott.ssa Lucia Allione"
    RAMM.email <- 'lucia.allione@ipsp.cnr.it'
    al.RSS <- "Al responsabile della sede secondaria di Legnaro"
    firma.RSS <- "Il responsabile della sede secondaria di Legnaro dell'IPSP"
    fatturazione <- "Istituto per la Protezione Sostenibile delle Piante, viale dell’Università, 16, 35020 Legnaro (PD), Italia."
    nomina.RSS <- "3905 dell'8/1/2025 di nomina della dott.ssa Laura Scarabel quale Responsabile della Sede Secondaria di Legnaro dell’IPSP per il periodo dall'8/1/2025 al 31/12/2025"
    nomina.RSS2 <- "3903 dell'8/1/2025 relativo al conferimento dell’incarico della dott.ssa Laura Scarabel a Responsabile delegato alla gestione della sede secondaria di Legnaro dell’IPSP con decorrenza dall'8/1/2025 al 31/12/2025;"
    nomina.RAMM <- "146196 del 2/5/2024 di nomina della dott.ssa Lucia Allione quale Responsabile Amministrativo della Sede Secondaria di Legnaro dell’IPSP per il periodo dal 1/5/2024 al 31/12/2024;"
    resp.segr <- "La responsabile amministrativa"
    sottoscritto.rss <- "La sottoscritta "
    nato.rss <- " nata a Bruxelles il 20/3/1963, codice fiscale SCRLRA63C60Z103Z,"
    RSS.dich <- " responsabile della sede secondaria di Legnaro dell'IPSP "
    CUU <- "8INQPI"
    cdr <- paste("CdR 121.00_.000 IPSP", sede2)
  }else if(sede=='TOsi'){
    sede1 <- 'Torino'
    sede2 <- 'Sede Istituzionale'
    RSS.nome <- 'Francesco Di Serio'
    RSS <- paste("dott.", RSS.nome)
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
    nato.rss <- " nato a Cava de' Tirreni (SA) il 29/09/1965, codice fiscale DSRFNC65P29C361R,"
    RSS.dich <- "direttore dell'IPSP"
    CUU <- "7K0RCK"
    cdr <- paste("CdR 121.000.000 IPSP", sede2)
  }

  if(Scelta.fornitore=='Avviso pubblico'){
    preventivo.individuato <- paste0("stato individuato per motivazioni tecnico-scientifiche e di economicità tra i preventivi di ",
                                     Prot..preventivi.avviso,
                                     " pervenuti in seguito all'avviso pubblico prot. ",
                                     Prot..avviso.pubblico,
                                     ";")
  #}else if(Scelta.fornitore=='Più preventivi'){
  }else{
    preventivo.individuato <- "individuato mediante indagine informale di mercato effettuata su MePA, mercato libero e/o cataloghi accessibili in rete, ritenuto in grado di assicurare la fornitura/il servizio richiesto secondo i tempi e le modalità indicati dall’Amministrazione, garantendo le migliori condizioni economiche e tecnico-qualitative;"
  #}else{
    #preventivo.individuato <- "allegato alla richiesta medesima e ritenuto in grado di assicurare la fornitura o la prestazione secondo i tempi e le modalità indicati dall’amministrazione, garantendo le migliori condizioni economiche e tecnico-qualitative;"
  }

  if(Rotazione.fornitore=="Importo <5.000€"){
    rotazione.individuata <- "che, in relazione a quanto indicato all'art. 49, comma 6, del Codice è possibile derogare dall'applicazione del principio di rotazione in caso di affidamenti di importo inferiore a euro 5.000,00;"
  }else if(Rotazione.fornitore=="Non è il contraente uscente"){
    rotazione.individuata <- "che in applicazione del principio di rotazione di cui all’art. 49, comma 2 del Codice, l’operatore economico individuato non è il contraente uscente;"
  }else if(Rotazione.fornitore=="Importo di fascia differente"){
    rotazione.individuata <- "che è possibile procedere all’affidamento al contraente uscente poiché non trova applicazione il principio di rotazione in quanto l’affidamento, pur riguardando l’operatore uscente, ha un importo appartenente ad una fascia diversa rispetto al precedente affidamento, in accordo con le linee guida per l’applicazione del principio di rotazione di affidamento dei contratti pubblici ai sensi del d.lgs. 36/2023 approvate dal CdA del CNR con delibera n. 412 del 19 dicembre 2023;"
  }else if(Rotazione.fornitore=="Particolare struttura del mercato"){
    rotazione.individuata <- "che, ai sensi dell’art. 49, comma 4, del Codice, è possibile procedere all’affidamento al contraente uscente in deroga al principio di rotazione in conseguenza della particolare struttura del mercato e dell'effettiva assenza di alternative e accertato che l'affidatario medesimo ha svolto accuratamente il precedente contratto garantendo altresì la qualità della prestazione resa;"
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
    fornitura.eseguita <- 'è stata consegnata'
  }else if(Tipo.acquisizione=='Servizi'){
    bene <- 'servizio'
    beni <- 'servizi'
    della.fornitura <- 'del servizio'
    la.fornitura <- 'il servizio'
    fornitura.consegnata <- 'il servizio dovrà essere prestato'
    materiale.conforme <- "che il servizio è conforme all’ordine e completamente prestato."
    fornitura.eseguita <- 'è stato eseguito'
  }else if(Tipo.acquisizione=='Lavori'){
    bene <- 'lavoro'
    beni <- 'lavori'
    della.fornitura <- 'del lavoro'
    la.fornitura <- 'il lavoro'
    fornitura.consegnata <- 'il lavoro dovrà essere svolto'
    materiale.conforme <- "che il lavoro è conforme all’ordine e completamente svolto."
    fornitura.eseguita <- 'è stato eseguito'
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
  if(RUP=="Maurizio Meoni" | RUP=="Giovanni Torraca" | RUP=="Salvatore Cristadoro"){
    Dott.rup <- 'Sig.'
    dott.rup <- 'sig.'
    il.dott.rup <- 'il sig.'
    al.rup <- 'al'
    dal.rup <- 'dal sig.'
    Nato.rup <- 'Nato a'
    nato.rup <- 'nato a'
    sottoscritto.rup <- 'Il sottoscritto'
    nominato <- "stato nominato"
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

  # if(length(anno==1)){
  # }else{
  #   anno <- grep("\\/2024$", sc$Prot..RAS)
  # }
  # 
  # #if(ordine>40 | ordine<1){
  # if(length(anno==1)){
  #   y <- "/2024"
  #   y2 <- 2024
  # }else{
  #   y <- "/2025"
  #   y2 <- 2025
  # }
  
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
  int.docoe <- toupper(paste0("Affidamento diretto, ai sensi dell’art. 50 del D.Lgs. N. 36/2023, ",
                            della.fornitura, " di “", Prodotto, "” (", Pagina.web, ") ",
                            "nell'ambito del progetto “", Progetto, "”", CUP1,
                            ", ordine ", sede, " ", ordine, y, ".", sep=""))

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
    finanziamento <- "No"
  }
  if(PNRR=="Agritech Spoke 1" | PNRR=="Agritech Spoke 3" | PNRR=="Agritech Spoke 8" | PNRR=="Agritech Spoke 4" | PNRR=="onFoods Spoke 2" | PNRR=="onFoods Spoke 3" | PNRR=="onFoods Spoke 4" | PNRR=="SUS-MIRRI.IT" | PNRR=="ITINERIS"){
    finanziamento <- "PNRR"
  }else if(PNRR=="DIVINGRAFT" | PNRR=="ARES" | PNRR=="MINACROP" | PNRR=="MONTANA" | PNRR=="SpecFor" | PNRR=="Mimic" | PNRR=="StreeTLAMP" | PNRR=="Fore-VOC" | PNRR=="XyWall" | PNRR=="AlpEcoArchaeology" | PNRR=="HEMINT" | PNRR=="BIORES" | PNRR=="Intertruffle" | PNRR=="BACBIO" | PNRR=="MOBeeFO" | PNRR=="secrEVome" | PNRR=="SUSHI" | PNRR=="TRSH52" | PNRR=="LICAT" | PNRR=="REMIND" | PNRR=="MYCOVIROLE"){
    finanziamento <- "PRIN 2022"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 104 del 2/2/2022 di emanazione del bando per i progetti di rilevante interesse nazionale (PRIN) 2022, nell’ambito del piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2, investimento 1.1;"
    mis.com.inv.esteso <- "piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.1 “progetti di ricerca di significativo interesse nazionale (PRIN)”, finanziamento dell'Unione europea - NextGeneration EU, decreto direttoriale MUR n. 104 del 2/2/2022"
    investimento <- "Investimento 1.1 “progetti di ricerca di rilevante interesse nazionale (PRIN)”"
  }else if(PNRR=="CIRCUFUN" | PNRR=="KNOWS" | PNRR=="PEP-HERB" | PNRR=="NEUROPLANT" | PNRR=="SAVEASH" | PNRR=="RNAi_Pj" | PNRR=="RE-VOC" | PNRR=="TOMRESAMED" | PNRR=="PLASMA4SOIL"){
    finanziamento <- "PRIN 2022 PNRR"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 1409 del 14/9/2022 di emanazione del bando per i progetti di rilevante interesse nazionale (PRIN) 2022, nell’ambito del piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2, investimento 1.1;"
    mis.com.inv.esteso <- "piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.1 “progetti di ricerca di significativo interesse nazionale (PRIN)”, finanziamento dell'Unione europea - NextGeneration EU, decreto direttoriale MUR n. 1409 del 14/09/2022"
    investimento <- "Investimento 1.1 “progetti di ricerca di rilevante interesse nazionale (PRIN)”"
  }
  
  if(PNRR=="Agritech Spoke 1" | PNRR=="Agritech Spoke 3" | PNRR=="Agritech Spoke 8" | PNRR=="Agritech Spoke 4"){
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
  if(PNRR=="Agritech Spoke 1"){
    Progetto.int <- sub("spoke 3", "spoke 1", Progetto.int)
  }
  if(PNRR=="Agritech Spoke 4"){
    Progetto.int <- sub("spoke 3", "spoke 4", Progetto.int)
  }
  if(PNRR=="Agritech Spoke 8"){
    Progetto.int <- sub("spoke 3", "spoke 8", Progetto.int)
    dicitura.fattura <- "Finanziamento Unione Europea NextGenerationEU progetto PNRR AGRITECH Spoke8 M4.C2.I1.4 - Codice progetto MUR: CN00000022"
  }
  if(PNRR=="onFoods Spoke 2" | PNRR=="onFoods Spoke 3" | PNRR=="onFoods Spoke 4"){
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
  if(PNRR=="onFoods Spoke 2"){
    Progetto.int <- sub("spoke 4", "spoke 2", Progetto.int)
  }
  if(PNRR=="onFoods Spoke 3"){
    Progetto.int <- sub("spoke 4", "spoke 3", Progetto.int)
  }
  if(PNRR=="SUS-MIRRI.IT"){
    Progetto.int <- 'piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 3.1 “fondo per la realizzazione di un sistema integrato di infrastrutture di ricerca e innovazione”, finanziato dall’Unione europea, NextGenerationEU, decreto direttoriale MUR n. 3264 del 28/12/2021 - progetto “Strengthening the MIRRI Italian Research Infrastructure for Sustainable Bioscience and Bioeconomy” (SUS-MIRRI.iT), codice identificativo IR0000005 - CUP D13C22001390001'
    codice.progetto <- "IR0000005"
    CUP2 <- "D13C22001390001"
    decreto.concessione <- "114 del 21/6/2022"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 3264 del 28/12/2021 di emanazione di un “avviso pubblico per la presentazione di proposte progettuali per “rafforzamento e creazione di infrastrutture di ricerca” da finanziare nell’ambito del Piano Nazionale di Ripresa e Resilienza, missione 4, componente 2, investimento 3.1 “fondo per la realizzazione di un sistema integrato di infrastrutture di ricerca e innovazione”, finanziato dall’Unione europea – NextGenerationEU;"
    decreto.pnrr <- " il Decreto Direttoriale MUR n. 114 del 21/6/2022 (di seguito, “Decreto di Concessione del Finanziamento”) col quale è stata ammessa a finanziamento la domanda di agevolazione presentata dall'Università di Torino, contrassegnata dal codice identificativo IR0000005, per la realizzazione del programma di ricerca dal titolo “Strengthening the MIRRI Italian Research Infrastructure for Sustainable Bioscience and Bioeconomy” (SUS-MIRRI.iT), CUP D13C22001390001;"
    dicitura.fattura <- "PNRR SUS-MIRRI.IT - Codice progetto MUR: IR0000005"
    investimento <- "3.1 “fondo per la realizzazione di un sistema integrato di infrastrutture di ricerca e innovazione”"
    intervento <- "SUS-MIRRI.IT - Codice progetto: IR0000005"
    mis.com.inv.esteso <- "piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 3.1 “fondo per la realizzazione di un sistema integrato di infrastrutture di ricerca e innovazione”, finanziamento dell'Unione europea - NextGeneration EU, decreto direttoriale MUR n. 3264 del 28/12/2021"
    attuatore <- "Università di Torino"
    avvio <- "Avvio: 1/11/2022; Conclusione: 31/10/2025"
    costo.totale <- "16.949.360,37 €"
    costo.ammesso <- "16.949.360,37 €"
    logo <- "logo_mirri.jpeg"
  }
  if(PNRR=="ITINERIS"){
    Progetto.int <- 'piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.3 “partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base”, finanziato dall’Unione europea, NextGenerationEU, decreto direttoriale MUR n. 130 del 21/06/2022 - progetto “Italian integrated environmental research infrastructures system” (ITINERIS), codice identificativo IR0000032 - CUP B53C22002150006'
    codice.progetto <- "IR0000032"
    CUP2 <- "B53C22002150006"
    decreto.concessione <- "130 del 21/6/2022"
    avviso.pnrr <- " il Decreto Direttoriale MUR n. 3264 del 28/12/2021 di emanazione di un “Avviso pubblico per la presentazione di Proposte di intervento per la creazione di “Partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base” nell’ambito del piano nazionale di ripresa e resilienza (PNRR), missione 4 “istruzione e ricerca”, componente 2 “dalla ricerca all’impresa”, investimento 1.3 “partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base”, finanziato dall’Unione europea, NextGenerationEU;"
    decreto.pnrr <- " il Decreto Direttoriale MUR n. 130 del 21/6/2022, registrato alla Corte dei Conti il 20/7/2022 al n. 1926 (di seguito, “Decreto di Concessione del Finanziamento”) col quale è stata ammessa a finanziamento la domanda di agevolazione presentata dal Consiglio Nazionale delle Ricerche, contrassegnata dal codice IR0000032, per la realizzazione del programma di ricerca dal titolo “Italian integrated environmental research infrastructures system” (ITINERIS), CUP B53C22002150006;"
    dicitura.fattura <- "PNRR ITINERIS - Codice progetto MUR: IR0000032"
    investimento <- "Investimento 1.3 “partenariati estesi alle università, ai centri di ricerca, alle aziende per il finanziamento di progetti di ricerca di base”"
    intervento <- "ITINERIS, codice progetto IR0000032"
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 1/11/2022; Conclusione: 30/4/2025 (prorogata al 31/10/2025)"
    costo.totale <- "____________ €"
    costo.ammesso <- "____________ €"
    logo <- "logo_itineris.png"
  }
  if(PNRR=="DIVINGRAFT"){
    titolo <- "dissection of molecular mechanisms underlying tolerance to virus and viroid infection in grafted tomato plants"
    codice.progetto <- "2022BZW9PF"
    CUP2 <- "B53D23017480006"
    decreto.concessione <- "1048 del 14/7/2023"
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
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 1409/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Ivan BACCELLI")
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
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 1409/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Federico BRILLI")
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
    dicitura.fattura <- paste0("Finanziamento Unione Europea NextGenerationEU, avviso 1409/2022 M4,C2,I1.1, codice ", codice.progetto, " “", PNRR, "”, CUP ", CUP2, ", resp. sci. Alberto SANTINI")
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "264.789,00 €, di cui a Unità CNR (IPSP-FI) 119.585,00 €"
    costo.ammesso <- "222.305,00 €, di cui a Unità CNR (IPSP-FI) 102.000,00 €"
    logo <- "logo_saveash.jpg"
  }
  if(PNRR=="XyWall"){
    titolo <- "cell wall determinants in plant resistance to Xylella"
    codice.progetto <- "2022F8BZMX"
    CUP2 <- "B53C24007440006"
    decreto.concessione <- "20427 del 6/11/2024"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 2/2/2025; Conclusione: 2/2/2027"
    costo.totale <- "284.374,00 €, di cui a Unità CNR (IPSP-BA) 132.429,00 €"
    costo.ammesso <- "203.387,00 €, di cui a Unità CNR (IPSP-BA) 68.342,00 €"
    logo <- "logo_xywall.tiff"
  }
  if(PNRR=="AlpEcoArchaeology"){
    titolo <- "West and East: an interdisciplinary approach to the archaeology of Alpine ecosystems"
    codice.progetto <- "2022T4T3Y8_SH6_PRIN2022"
    CUP2 <- "B53D23001940006"
    decreto.concessione <- "SH6 _MUR n. 969 del 30/06/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Verona, Prof. MIGLIAVACCA Maria Gioia"
    avvio <- "Avvio: 28/9/2023; Conclusione: 27/9/2025"
    costo.totale <- "302.236,00 €; quota del CNR-IPSP: 53.342,00 € = 17.940,00 € cofin. + 35.402,00 € contributo MUR"
    costo.ammesso <- "302.236,00 €"
    logo <- "logo_alpeco.jpg"
  }
  if(PNRR=="HEMINT"){
    titolo <- "RNA interference for the control of hemipteran pests causing direct and indirect damages to crops: a case study to define best practices for sustainable and environment-friendly application of RNAi (HEMipteran INTerference)"
    codice.progetto <- "2022BPB5A8"
    CUP2 <- "B53D23017470006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Torino"
    avvio <- "Avvio: 11/10/2023; Conclusione: 10/10/2025"
    costo.totale <- "283.646,00 €, di cui 91.670,00 € al CNR-IPSP"
    costo.ammesso <- "234.858,00 €, di cui 75.902,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="BIORES"){
    titolo <- "Natural and BIOtechnological genetic RESistances against Flavescence dorée for vineyard sustainability"
    codice.progetto <- "2022FW39MT"
    CUP2 <- "B53D23017600006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 11/10/2023; Conclusione: 10/10/2025"
    costo.totale <- "311.065,00 €, di cui 166.210,00 € al CNR-IPSP"
    costo.ammesso <- "257.353,00 €, di cui 133.867,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="RNAi_Pj"){
    titolo <- "RNA interference to control alien pests by exploring microalgae and cyanobacteria as cost-effective dsRNA producing platforms: the Japanese beetle Popillia japonica as case study"
    codice.progetto <- "P20227835Y"
    CUP2 <- "B53D23031970001"
    decreto.concessione <- "1377 del 01/09/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Verona"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "299.952,00 €, di cui 135.474,00 € al CNR-IPSP"
    costo.ammesso <- "224.965,00 €, di cui 99.800,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="Intertruffle"){
    titolo <- "Interactions of the white truffle Tuber magnatum with soil microbiome and plant"
    codice.progetto <- "2022K272X8_LS9_PRIN2022"
    CUP2 <- "B53D23017740006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università Telematica San Raffaele Roma"
    avvio <- "Avvio: 12/10/2023; Conclusione: 11/10/2025"
    costo.totale <- "303.444,00 €, di cui 39.734 € al CNR-IPSP"
    costo.ammesso <- "303.444,00 €, di cui 39.734 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="BACBIO"){
    titolo <- "Molecular genetics and genomics of fruit fly associated bacteria for implementation of innovative biocontrol strategies"
    codice.progetto <- "2022LEW75T"
    CUP2 <- "B53C24007460006"
    decreto.concessione <- "20427 del 6/11/2024"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 6/2/2025; Conclusione: 5/2/2027"
    costo.totale <- "248.465,00 €, di cui 167.465,00 € al CNR-IPSP"
    costo.ammesso <- "203.272,00 €, di cui 122.272,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="MOBeeFO"){
    titolo <- "Monitoring of honey Bee immunomodulation and resilience to stress factors by Fiber Optic technology"
    codice.progetto <- "2022YRJAAC"
    CUP2 <- "B53D23003030006"
    decreto.concessione <- "960 del 30/6/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università di Napoli Parthenope"
    avvio <- "Avvio: 1/10/2023; Conclusione: 30/9/2025"
    costo.totale <- "232.954,00 €, di cui 77.968,00 € al CNR-IPSP"
    costo.ammesso <- "232.954,00 €, di cui 77.968,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="RE-VOC"){
    titolo <- "Looking for plant VOC receptors"
    codice.progetto <- "P20229ZW4A"
    CUP2 <- "B53D23032030001"
    decreto.concessione <- "1377 dell'1/9/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "299.862,00 €, di cui 150.053,00 € al CNR-IPSP"
    costo.ammesso <- "224.896,00 €, di cui 112.539,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="secrEVome"){
    titolo <- "The complex role of plant extracellular vesicles: deciphering their secreted molecular messages and bioactivity in plant-microbe"
    codice.progetto <- "202224M943"
    CUP2 <- "B53D23017030006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Salerno"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "291.500,00 €"
    costo.ammesso <- "241.562,00 €"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="SUSHI"){
    titolo <- "SUccess of Specialist versus generalist parasitoid in Hampering the spread of an Invasive pest"
    codice.progetto <- "_______"
    CUP2 <- "___"
    decreto.concessione <- "739 del 29/5/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 1/10/2023; Conclusione: 30/9/2025"
    costo.totale <- "264.369,60 €, di cui 82.877,20 € al CNR-IPSP"
    costo.ammesso <- "205.637,89 €, di cui 63.152,20 € al CNR-IPSP"
    logo <- "logo_sushi.jpg"
  }
  if(PNRR=="TOMRESAMED"){
    titolo <- "Tomato genetic diversity for enhancing Resilience of Agro-systems in Mediterranean environment"
    codice.progetto <- "P2022LP2YW"
    CUP2 <- "B53D23023640001"
    decreto.concessione <- "1180 del 27/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università degli Studi di Bari Aldo Moro"
    avvio <- "Avvio: 9/11/2023; Conclusione: 8/11/2025"
    costo.totale <- "300.000,00 €, di cui 60.000,00 € al CNR-IPSP"
    costo.ammesso <- "237.300,00 €, di cui 47.460,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="TRSH52"){
    titolo <- "A supplementary diet as therapeutic vaccine for pancreatic cancer"
    codice.progetto <- "2022TRSH52_LS6"
    CUP2 <- "B53D23003580006"
    decreto.concessione <- "972 del 30/6/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università di Torino"
    avvio <- "Avvio: 1/11/2023; Conclusione: 30/10/2025"
    costo.totale <- "284.024,00 €, di cui 112.824,00 € al CNR-IPSP"
    costo.ammesso <- "211.209,00 €, di cui 105.604,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="LICAT"){
    titolo <- "LISCL-mediated catabolism regulates plant-environment interactions"
    codice.progetto <- "2022BKBMLM_LS2_PRIN2022"
    CUP2 <- "B53D23007960006"
    decreto.concessione <- "752 dell'1/6/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Università di Torino"
    avvio <- "Avvio: 28/9/2023; Conclusione: 27/9/2025"
    costo.totale <- "249.879,77 €, di cui 83.302,77 € al CNR-IPSP"
    costo.ammesso <- "224.732,00 €, di cui 69.864,00 € al CNR-IPSP"
    logo <- "logo_bianco.jpg"
  }
  if(PNRR=="REMIND"){
    titolo <- "Do crop plants remember stress? Effect of water stress memory on crop resilience in response to recurrent drought and recovery events"
    codice.progetto <- "2022RBHRJR"
    CUP2 <- "B53D23018020006"
    decreto.concessione <- "1048 del 14/7/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 12/10/2023; Conclusione: 28/2/2026 (come da proroga DD MUR n. 509 del 16/4/2025"
    costo.totale <- "276.104,00 €, di cui 110.386,00 € al CNR-IPSP"
    costo.ammesso <- "206.191,00 €, di cui 87.303,00 € al CNR-IPSP"
    logo <- "logo_remind.tif"
  }
  if(PNRR=="MYCOVIROLE"){
    titolo <- "Viromes of fungi interacting with plants: roles in adaptation to different/extreme ecological niches and reservoir of viral biodiversity for potential spillover"
    codice.progetto <- "20222L5ECJ_LS8_PRIN2022"
    CUP2 <- "B53D23011710006"
    decreto.concessione <- "739 del 29/5/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche"
    avvio <- "Avvio: 4/10/2023; Conclusione: 3/10/2025"
    costo.totale <- "306.638,00 €, di cui 99.194,00 € al CNR-IPSP"
    costo.ammesso <- "252.364,00 €, di cui 81.638,00 € al CNR-IPSP"
    logo <- "logo_mycovirole.png"
  }
  if(PNRR=="PLASMA4SOIL"){
    titolo <- "Development of essential oil-based smart formulates by means of plasma processing: effect against pests and impact on soil beneficial communities"
    codice.progetto <- "P2022MK3AF"
    CUP2 <- "B53D23032200001"
    decreto.concessione <- "1289 del 4/8/2023"
    dicitura.fattura <- paste0(finanziamento, " ", PNRR, " - Codice progetto MUR: ", codice.progetto)
    attuatore <- "Consiglio Nazionale delle Ricerche, Nanotec"
    avvio <- "Avvio: 30/11/2023; Conclusione: 29/11/2025"
    costo.totale <- "300.000,00 €"
    costo.ammesso <- "225.000,00 €"
    logo <- "logo_plasma4soil.jpg"
  }
  
  dicitura.fatturazione <- paste0("Si prega di riportare in fattura le seguenti informazioni: ordine n° ", sede, " ", ordine, y, ", prot. n. _____ (si veda in alto nella pagina della lettera d'ordine), CIG ", CIG, ", CUP ", CUP, ".")
  dicitura.fatturazione.eng <- paste0("In the invoice, plese report the following information: purchase order n° ", sede, " ", ordine, y, ", prot. n. _____ (see on the top of the purchase order page), CIG ", CIG, ", CUP ", CUP, ".")
  
  if(PNRR!="No"){
    dicitura.fatturazione <- sub(".$", paste0(", progetto '", dicitura.fattura, "'."), dicitura.fatturazione)
    dicitura.fatturazione.eng <- sub(".$", paste0(", project '", dicitura.fattura, "'."), dicitura.fatturazione.eng)
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
    Progetto.int.no.cup <- sub(" .CUP.*", "", Progetto.cup)
  }else{
    Progetto.int.no.cup <- sub(" - CUP .*", "", Progetto.int)
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
  fornitore.uscente <- "falso"
  blocco.rota <- "falso"
  
  if(CPV!=trattini){
    rota <- subset(ordini, substr(ordini$CPV, 1, 3)==substr(sc$CPV, 1, 3))
    rota$Data <- as.POSIXct(rota$Data, tz="CET", format = "%d/%m/%Y")
    rota$Fascia <- ifelse(rota$Importo.senza.IVA.num<5000, "< 5.000 €", 
                          ifelse(rota$Importo.senza.IVA.num>=40000, "> 40.000 €", "5.000 - 40.000 €" )                        )
    rota <- dplyr::select(rota, Ordine.N., Data, Fornitore, CPV, Prodotto, Importo.senza.IVA, Importo.senza.IVA.num, Fascia, Rotazione.fornitore)
    rota <- subset(rota, !is.na(rota$Fornitore) & !is.na(rota$CPV) & !is.na(rota$Importo.senza.IVA.num))
    rownames(rota) <- NULL
    s <- subset(rota, rota$Ordine.N.==ordine.orig)
    n <- as.numeric(rownames(s))
    #n <- grep(ordine.orig, rota$Ordine.N., fixed = TRUE) 
    if(n>1){rota <- rota[-1:-(n-1),]}
    if(length(rota$Fornitore)>1 & rota$Fornitore[1] == rota$Fornitore[2]){fornitore.uscente <- "vero"}
    if(length(rota$Fornitore)>1 & rota$Fascia[1] == rota$Fascia[2]){fascia <- "stessa"}else{fascia <- "diversa"}
    frase1 <- frase2 <- ""
    frase3.1 <- "   E' possibile derogare alla rotazione dei fornitori e, quindi, affidare a questo fornitore "
    frase3.3 <- " nella colonna 'Rotazione fornitore' di FluOr"
    frase3.4 <- ". Assicurarsi che questa scelta sia descritta nella relazione della Richiesta d'Acquisto.\n"
    frase4 <- "*********************\n\n"
  }else{
    stop("CPV mancante! Inserire il CPV in FluOr, scaricare nuovamente Ordini.csv e generare i documenti.")
  }
  
  if(fornitore.uscente=="vero"){
    frase1 <- paste0("\n***** ATTENZIONE *****\n   ", Fornitore, " è il fornitore uscente.\n",
                     "   L'ultimo ordine (n° ", rota$Ordine.N.[2], ") per questa categoria merceologica (prime tre cifre del CPV: ", substr(sc$CPV, 1, 3), ") è stato affidato a questo operatore economico per l'acquisto di '", rota$Prodotto[2], "' e un importo di € ", rota$Importo.senza.IVA[2], ".\n")
    if(fascia=="stessa"){
      frase2 <- paste0("   L'ordine in oggetto (n° ", ordine, ") e l'ordine n° ", rota$Ordine.N.[2], " hanno la stessa fascia d'importo (", rota$Fascia[1], ").\n")
      if(Importo.senza.IVA.num<5000){
        if(Rotazione.fornitore=="Importo <5.000€"){
          frase3.2 <- "poichè si è indicato 'Importo <5.000€'"
        }else if(Rotazione.fornitore=="Particolare struttura del mercato"){
          frase3.2 <- "poichè si è indicato 'Particolare struttura del mercato'"
        }else{
          frase3.2 <- "impostando 'Importo <5.000€' oppure 'Particolare struttura del mercato'"
          frase3.4 <- " e motivando questa scelta nella relazione della Richiesta d'Acquisto.\n"
          blocco.rota <- "vero"
        }
      }else{
        if(Rotazione.fornitore=="Particolare struttura del mercato"){
          frase3.2 <- "poiché si è indicato 'Particolare struttura del mercato'"
        }else{
          frase3.2 <- "impostando 'Particolare struttura del mercato'"
          frase3.4 <- " e motivando questa scelta nella relazione della Richiesta d'Acquisto.\n"
          blocco.rota <- "vero"
        }
      }
    }else if(fascia=="diversa"){
      frase2 <- paste0("   L'ordine in oggetto (n° ", ordine, "; ", rota$Fascia[1], ") e l'ordine n° ", rota$Ordine.N.[2], " (", rota$Fascia[2], ") hanno fascia d'importo differente.\n")
      if(Importo.senza.IVA.num<5000){
        if(Rotazione.fornitore=="Importo <5.000€"){
          frase3.2 <- " poichè si è indicato 'Importo <5.000€'"
        }else if(Rotazione.fornitore=="Particolare struttura del mercato"){
          frase3.2 <- "poichè si è indicato 'Particolare struttura del mercato'"
        }else if(Rotazione.fornitore=="Importo di fascia differente"){
          frase3.2 <- "poichè si è indicato 'Importo di fascia differente'"
        }else{
          frase3.2 <- "impostando 'Importo di fascia differente', 'Importo <5.000€' oppure 'Particolare struttura del mercato'"
          frase3.4 <- " e motivando questa scelta nella relazione della Richiesta d'Acquisto.\n"
          blocco.rota <- "vero"
        }
      }else if(Importo.senza.IVA.num>=5000){
        if(Rotazione.fornitore=="Particolare struttura del mercato"){
          frase3.2 <- "poiché si è indicato 'Particolare struttura del mercato'"
        }else if(Rotazione.fornitore=="Importo di fascia differente"){
          frase3.2 <- "poichè si è indicato 'Importo di fascia differente'"
        }else{
          frase3.2 <- "impostando 'Importo di fascia differente' oppure 'Particolare struttura del mercato'"
          frase3.4 <- " e motivando questa scelta nella relazione della Richiesta d'Acquisto.\n"
          blocco.rota <- "vero"
        }
      }
    }
  }

  # vecchia Rotazione fornitore ----
  # if(sc$Importo.senza.IVA.num<5000){
  #   ordini.fascia <- subset(ordini, ordini$Importo.senza.IVA.num<5000)
  #   }else if(sc$Importo.senza.IVA.num>=5000 & sc$Importo.senza.IVA.num<40000){
  #     ordini.fascia <- subset(ordini, ordini$Importo.senza.IVA.num>=5000 & ordini$Importo.senza.IVA.num<40000)
  #   }else if(sc$Importo.senza.IVA.num>=40000){
  #     ordini.fascia <- subset(ordini, ordini$Importo.senza.IVA.num>=40000)
  #   }
  # 
  # rota <- subset(ordini.fascia, ordini.fascia$CPV==sc$CPV)
  # rota <- dplyr::select(rota, Ordine.N., Data, Fornitore, CPV, Prodotto, Importo.senza.IVA)
  # rota$Data <- as.POSIXct(rota$Data, tz="CET", format = "%d/%m/%Y")
  # data.ordine <- subset(rota, rota$Ordine.N.==ordine)
  # data.ordine <- data.ordine$Data
  # rota <- subset(rota, rota$Data<=data.ordine)
  # rota$CPV.iniz <- sub("(...).*", "\\1", rota$CPV)
  # rota <- subset(rota, rota$Ordine.N.!=ordine)
  # rota <- rota[order(rota$Data),]
  # lng <- length(rota$Ordine.N.)
  # rota <- rota[lng,]
  # rota.display <- dplyr::select(rota, Ordine.N., Fornitore, CPV, Prodotto, Importo.senza.IVA)
  # ordine.uscente <- rota$Ordine.N.
  # fornitore.uscente <- rota$Fornitore
  # cpv.usente <- rota$CPV.iniz
  # prodotto.uscente <- rota$Prodotto
  # importo.uscente <- rota$Importo.senza.IVA
  # if(lng==0){fornitore.uscente <- "nessuno"}

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
  doc.dic.pres <- read_docx("Modello.docx")
  download.file("https://raw.githubusercontent.com/giovabubi/appost/main/models/Modello_intestata.docx", destfile = "Modello.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
  doc.prov.liq <- read_docx("Modello.docx")
  file.remove("Modello.docx")

  # RAS ----
  ras <- function(){
    if(fornitore.uscente=="vero"){
      cat(paste0(frase1, frase2, frase3.1, frase3.2, frase3.3, frase3.4, frase4))
      if(blocco.rota=="vero"){
        stop("Non è possibile continuare. Apportare le modifiche in FluOr come indicato sopra e, poi, generare nuovamente i documenti dopo aver scaricato Ordini.csv.\n")
      }else{
        cat("E' possibile continuare. Premere INVIO per proseguire\n")
        readline()
      }
    }
    
    # if(Fornitore==fornitore.uscente){
    #   cat(paste0(
    #     "***** ATTENZIONE *****\n",
    #     Fornitore, " è il fornitore uscente.\n",
    #     "L'ultimo ordine (n° ", ordine.uscente, ") per questa categoria merceologica (prime tre cifre del CPV: ", cpv.usente, ") è stato affidato a questo operatore economico per l'acquisto di '", prodotto.uscente, "' e un importo di € ", importo.uscente, ".\n"))
    #   if(Rotazione.fornitore=="Non è il contraente uscente"){
    #     cat("In FluOr è stato erroneamente indicato 'Non è il contraente uscente'. Si prega di apportare la dovuta correzione.\n")
    #   }else if(Rotazione.fornitore=="Particolare struttura del mercato"){
    #     cat("L'ordine può procedere poichè è stato indicato 'Particolare struttura del mercato'.\n")
    #   }else if(Rotazione.fornitore=="Importo <5.000€" & Importo.senza.IVA.num<5000){
    #     cat("L'ordine può procedere poichè è stata specificata la deroga alla rotazione dei fornitori per ordini <5.000 €.\n")
    #   }else if(Rotazione.fornitore=="Importo <5.000€" & Importo.senza.IVA.num>=5000){
    #     cat("E' stata specificata la deroga alla rotazione dei fornitori per ordini <5.000 €, ma l'ordine è superiore a questo importo. Si prega di apportare la dovuta correzione.\n")
    #   }
    #     cat("*********************\n",
    #     " Premere INVIO per proseguire")
    #   readline()
    # }

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
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1) |>
      headers_replace_text_at_bkm("bookmark_headers_sede2", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale") |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale2", "Istituzionale")
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
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2")
    
    if(Richiedente!=Responsabile.progetto){
      doc <- doc |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext(Dott.resp), ftext(" "), ftext(Responsabile.progetto)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(responsabile del progetto e titolare dei fondi)")), style = "Firma 2")
    }

    doc <- doc |>
      cursor_bookmark("bookmark_relazione") |>
      body_remove() |>
        cursor_backward()
      
    if(CPV..CPV=="22120000-7"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Una ricerca condotta nell’ambito delle attività del progetto "),
                           ftext(Progetto),
                           ftext(" è stata completata e convogliata in un articolo scientifico scritto da __________ e intitolato “__________”.")), style = "Relazione") |>
        body_add_fpar(fpar(ftext("Indagine di mercato")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("Un’indagine delle riviste scientifiche più adatte a questo articolo per i temi trattati e che abbiamo elevato impact factor, rientrino nel primo quartile (Q1) nel settore "),
                           ftext("Plant Sciences", fpt.i),
                           ftext(" e abbiano prezzi competitivi per pubblicazioni Open Access CC-BY ha portato all’individuazione della rivista __________ pubblicata da "),
                           ftext(Fornitore),
                           ftext(".")), style = "Relazione") |>
        body_add_fpar(fpar(ftext("Dopo peer review, l’articolo è stato ora accettato per la pubblicazione.")), style = "Relazione") |>
        body_add_fpar(fpar(ftext("L’operatore economico è, quindi, "),
                           ftext(Fornitore),
                           ftext(", che offre il servizio di pubblicazione Open Access al costo di "),
                           ftext(Importo.senza.IVA),
                           ftext(" IVA esclusa. Tale fornitore risulta, inoltre, in possesso di esperienze pregresse con altre pubbliche amministrazioni italiane ed è iscritto al MePA.")), style = "Relazione")
    
      if(Rotazione.fornitore!="Non è il contraente uscente"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("L’operatore economico individuato risulta essere contraente uscente. Tuttavia, si chiede l’affidamento all’operatore economico individuato in deroga al principio di rotazione per le seguenti motivazioni, ai sensi dell'art. 49, comma 4 del Codice:")), style = "Relazione") |>
          body_add_fpar(fpar(ftext("struttura del mercato, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("effettiva assenza di alternative, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("accurata esecuzione del precedente contratto, quale __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("trattasi di beni specifici prodotti esclusivamente dal fornitore individuato e funzionali all’attività di ricerca, che richiede continuità e ripetibilità di protocolli operativi specifici;")), style = "Elenco punto")
        if(Importo.senza.IVA.num<5000){
          doc <- doc |>
            body_add_fpar(fpar(ftext("l’importo dell’affidamento è inferiore a euro 5.000,00 (ai sensi dell’art. 49, comma 6, del Codice).")), style = "Elenco punto")
        }
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("Conclusioni")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("In seguito all’accettazione dell’articolo per la pubblicazione, sul sito della rivista viene mostrato il costo, che è pari a "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA. Si richiede, pertanto, l’attivazione dell’idoneo procedimento finalizzato all’acquisizione del servizio in oggetto.")), style = "Relazione")
    }else{
      if(Tipo.acquisizione=="Beni"){
        doc <- doc |>
        body_add_fpar(fpar(ftext("Per le attività di ricerca previste nel progetto "),
                           ftext(Progetto),
                           ftext(" è necessaria l'acquisizione "),
                           ftext(della.fornitura),
                           ftext(" di “"),
                           ftext(Prodotto),
                           ftext("”, come dettagliato nella richiesta d'acquisto.")), style = "Relazione")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("Per le attività di ricerca previste nel progetto "),
                             ftext(Progetto),
                             ftext(" è necessaria l'acquisizione "),
                             ftext(della.fornitura),
                             ftext(" di “"),
                             ftext(Prodotto),
                             ftext("” con le seguenti caratteristiche:")), style = "Relazione") |>
          body_add_fpar(fpar(ftext("__________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("__________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("__________;")), style = "Elenco punto")
      }
    doc <- doc |>
      body_add_fpar(fpar(ftext("Indagine di mercato")), style = "heading 2")
      
    if(Scelta.fornitore=="Singolo preventivo"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("In seguito ad un’accurata valutazione del mercato è stato acquisito un singolo preventivo, allegato alla presente, per la seguente motivazione: ___________.")), style = "Relazione")
          }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("In seguito ad un’accurata indagine informale di mercato, con la quale sono stati acquisiti n° ___ preventivi, allegati alla presente, è stato individuato l’operatore economico "),
                             ftext(Fornitore),
                             ftext(" quale potenziale affidatario "),
                             ftext(della.fornitura),
                             ftext(" per le seguenti motivazioni: ___________.")), style = "Relazione")
          }
      doc <- doc |>
        body_add_fpar(fpar(ftext("L’operatore economico "),
                           ftext(Fornitore),
                           ftext(" ci ha inviato un preventivo rispondente esattamente alle nostre richieste ed esigenze sia dal punto di vista delle caratteristiche tecniche che dei tempi di consegna, che dal punto di vista del prezzo rispondente agli standard di mercato e con tutte le garanzie richieste sui prodotti. Tale fornitore risulta inoltre in possesso delle esperienze pregresse idonee all’esecuzione della prestazione contrattuale, quali altre forniture simili a pubbliche amministrazioni compreso il CNR.")), style = "Relazione")
      if(Rotazione.fornitore!="Non è il contraente uscente"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("L’operatore economico individuato risulta essere contraente uscente. Tuttavia, si chiede l’affidamento all’operatore economico individuato in deroga al principio di rotazione per le seguenti motivazioni, ai sensi dell'art. 49, comma 4 del Codice:")), style = "Relazione") |>
          body_add_fpar(fpar(ftext("struttura del mercato, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("effettiva assenza di alternative, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("accurata esecuzione del precedente contratto, quale __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("trattasi di beni specifici prodotti esclusivamente dal fornitore individuato e funzionali all’attività di ricerca, che richiede continuità e ripetibilità di protocolli operativi specifici;")), style = "Elenco punto")
        if(Importo.senza.IVA.num<5000){
          doc <- doc |>
            body_add_fpar(fpar(ftext("l’importo dell’affidamento è inferiore a euro 5.000,00 (ai sensi dell’art. 49, comma 6, del Codice).")), style = "Elenco punto")
        }
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("Conclusioni")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("Da contatti informali, cui è seguita una quotazione budgetaria, il costo massimo omnicomprensivo atteso per l’acquisizione è pari a "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA. Si richiede, pertanto, l’attivazione dell’idoneo procedimento finalizzato all’acquisizione "),
                           ftext(della.fornitura),
                           ftext(" in oggetto.")), style = "Relazione")
    }
    doc <- doc |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da))) |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext(Dott.ric), ftext(" "), ftext(Richiedente)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2")
      
      if(Richiedente!=Responsabile.progetto){
        doc <- doc |>
          body_add_par("") |>
          body_add_par("") |>
          body_add_par("") |>
          body_add_fpar(fpar(ftext(Dott.resp), ftext(" "), ftext(Responsabile.progetto)), style = "Firma 2") |>
          body_add_fpar(fpar(ftext("(responsabile del progetto e titolare dei fondi)")), style = "Firma 2")
      }
      
    print(doc, target = paste0(pre.nome.file, "1 RAS.docx"))
    cat("

    Documento generato: '1 RAS'")

    ## Dich. Ass. Rich. ----
    download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50, COMMA 1, LETT. B) DEL D.LGS. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" DI “"),
                         ftext(Prodotto),
                         ftext("” nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", ordine "),
                         ftext(sede),
                         ftext(" "),
                         ftext(ordine),
                         ftext(y)), style = "Maiuscolo") |>
      body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.ric), ftext(" "), ftext(dott.ric), ftext(" "), ftext(Richiedente, fpt.b), ftext(", "),
                         ftext(nato.ric), ftext(" "), ftext(Richiedente..Luogo.di.nascita), ftext(", il "),
                         ftext(Richiedente..Data.di.nascita), ftext(", codice fiscale "), ftext(Richiedente..Codice.fiscale), ftext(",")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di "),
                         ftext("richiedente", fpt.b),
                         ftext(" la fornitura di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(" "),
                         ftext(all.OE),
                         ftext(", nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
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
      body_add_fpar(fpar(ftext(sede1), ftext(", "),ftext(da)), style = "Normal") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(paste0(Dott.ric," ", Richiedente), run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "9.1 Dichiarazione assenza conflitto RICH.docx"))
    
    cat("

    Documento generato: '9.1 Dichiarazione assenza conflitto RICH'")

    ## Dich. Ass. Resp. ----
    if(Richiedente!=Responsabile.progetto){
      download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
      
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
      }
      
      doc <- doc |>
        cursor_begin() |>
        body_add_fpar(fpar(ftext("All’"),
                           ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
        body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50, COMMA 1, LETT. B) DEL D.LGS. N. 36/2023, "),
                           ftext(della.fornitura), ftext(" DI “"),
                           ftext(Prodotto),
                           ftext("” nell'ambito del progetto “"),
                           ftext(Progetto),
                           ftext("”"),
                           ftext(CUP1),
                           ftext(", ordine "),
                           ftext(sede),
                           ftext(" "),
                           ftext(ordine),
                           ftext(y)), style = "Maiuscolo") |>
        body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext(sottoscritto.resp), ftext(" "), ftext(dott.resp), ftext(" "), ftext(Responsabile.progetto, fpt.b), ftext(", "),
                           ftext(nato.resp), ftext(" "), ftext(Responsabile.progetto..Luogo.di.nascita), ftext(", il "),
                           ftext(Responsabile.progetto..Data.di.nascita), ftext(", codice fiscale "), ftext(Responsabile.progetto..Codice.fiscale), ftext(",")), style = "Normal") |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b),
                           ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in qualità di titolare dei fondi e "),
                           ftext("responsabile del progetto di ricerca ", fpt.b),
                           ftext("“"),
                           ftext(Progetto), ftext("”"), ftext(CUP1),
                           ftext(", in relazione alla fornitura di “"),
                           ftext(Prodotto, fpt.b),
                           ftext("”, ordine "),
                           ftext(sede, fpt.b),
                           ftext(" "),
                           ftext(ordine, fpt.b),
                           ftext(y, fpt.b),
                           ftext(" all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext(");")), style = "Normal") |>
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
        body_add_fpar(fpar(ftext(sede1), ftext(", "),ftext(da)), style = "Normal") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(paste0(Dott.resp," ",Responsabile.progetto), run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(Responsabile del progetto e titolare dei fondi)")), style = "Firma 2")
      
      print(doc, target = paste0(pre.nome.file, "9.2 Dichiarazione assenza conflitto RESP.docx"))
      
      cat("

    Documento generato: '9.2 Dichiarazione assenza conflitto RESP'")

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore)
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
    
    ## Avviso pubblico ----
    if(Scelta.fornitore=='Avviso pubblico'){
      download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
      }
      doc <- doc |>
        cursor_begin() |>
        cursor_forward() |>
        #headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE) |>
        
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
        body_add_fpar(fpar(ftext("La Stazione Appaltante ISTITUTO PER LA PROTEZIONE SOSTENIBILE DELLE PIANTE del CNR intende procedere, a mezzo della presente indagine esplorativa, all’individuazione di un operatore economico a cui affidare eventualmente la fornitura/il servizio di cui all’oggetto, ai sensi dell’art. 50, comma 1 del d.lgs. 36/2023.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Il presente avviso, predisposto nel rispetto dei principi di libera concorrenza, non discriminazione, trasparenza, proporzionalità e pubblicità, non costituisce invito a partecipare a gara pubblica, né un’offerta al pubblico (art. 1336 del Codice civile) o promessa al pubblico (art. 1989 del Codice civile), ma ha lo scopo di esplorare le possibilità offerte dal mercato al fine di affidare direttamente "),
                           ftext(la.fornitura),
                           ftext(".")), style = "Normal") |>
        body_add_fpar(fpar(ftext("L’indagine in oggetto non comporta l’instaurazione di posizioni giuridiche ovvero obblighi negoziali. Il presente avviso, pertanto, non vincola in alcun modo questa Stazione Appaltante che si riserva, comunque, la facoltà di sospenderlo, modificarlo o annullarlo e di non dar seguito al successivo affidamento, senza che gli operatori economici possano vantare alcuna pretesa.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("I preventivi ricevuti si intenderanno impegnativi per gli operatori economici per un periodo di massimo di 60 giorni naturali e consecutivi, mentre non saranno in alcun modo impegnativi per la Stazione Appaltante, per la quale resta salva la facoltà di procedere o meno a successive e ulteriori richieste di preventivi volte all’affidamento "),
                           ftext(della.fornitura),
                           ftext(" di cui all’oggetto.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("L’affidamento sarà espletato attraverso una piattaforma di approvvigionamento digitale certificata.")), style = "Normal") |>
        
        body_add_fpar(fpar(ftext("Oggetto "), ftext(della.fornitura)), style = "heading 3") |>
        body_add_fpar(fpar(ftext("L’oggetto "), ftext(della.fornitura), ftext(" è _____.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("La consegna dovrà avvenire presso _____ entro _____.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("[Specificare tutte le caratteristiche del bene/servizio/lavoro, nonchè modalità e tempi di consegna, così che gli operatori economici possano presentare offerte comparabili e la stazione appaltante possa scegliere il preventivo più adatto in base ai criteri richiesti in fase di avviso pubblico]", fpt.i)), style = "Normal") |>
        
        body_add_fpar(fpar(ftext("Requisiti")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Possono inviare il proprio preventivo gli operatori economici in possesso di:")), style = "Normal") |>
        body_add_fpar(fpar(ftext("abilitazione MePA relativa al bando “"),
                                 ftext(beni),
                                 ftext("”, categoria “__________”;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("requisiti di ordine generale di cui al Capo II, Titolo IV del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("requisiti d’idoneità professionale come specificato all’art. 100, comma 3 del D.lgs. n. 36/2023: iscrizione nel registro della camera di commercio, industria, artigianato e agricoltura o nel registro delle commissioni provinciali per l’artigianato o presso i competenti ordini professionali per un’attività pertinente anche se non coincidente con l’oggetto dell’affidamento. All’operatore economico di altro Stato membro non residente in Italia è richiesto di dichiarare ai sensi del testo unico delle disposizioni legislative e regolamentari in materia di documentazione amministrativa, di cui al decreto del Presidente della Repubblica del 28 dicembre 2000, n. 445;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("documentate esperienze pregresse idonee all’esecuzione delle prestazioni contrattuali oggetto dell’affidamento.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("[eventuale]", fpt.i), ftext(" requisiti di capacità economico-finanziaria e/o tecnico-professionale.")), style = "Elenco punto") |>
        
        body_add_fpar(fpar(ftext("Valore dell'affidamento")), style = "heading 3")
      if(CCNL=="Non applicabile"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("La Stazione Appaltante ha stimato per l’affidamento di cui all’oggetto un importo massimo pari a "),
                             ftext(Importo.senza.IVA, fpt.b), ftext(" oltre IVA e/o altre imposte e contributi di legge.")), style = "Normal")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("La Stazione Appaltante ha stimato per l’affidamento di cui all’oggetto un importo massimo pari a "),
                             ftext(Importo.senza.IVA, fpt.b), ftext(" oltre IVA e/o altre imposte e contributi di legge, comprensivo di "),
                             ftext(Oneri.sicurezza),
                             ftext(" quali oneri per la sicurezza dovuti a rischi da interferenze e comprensivo di "),
                             ftext(Manodopera),
                             ftext(" quale importo totale dei costi della manodopera calcolato considerando il seguente CCNL territoriale: "),
                             ftext(CCNL),
                             ftext(".")), style = "Normal")
      }
      
      doc <- doc |>
        body_add_fpar(fpar(ftext("Modalità di presentazione del preventivo")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Gli operatori economici in possesso dei requisiti sopra indicati potranno inviare il proprio preventivo, corredato della dichiarazione attestante il possesso dei requisiti predisposta secondo il modello allegato al presente avviso (allegato 1), entro e non oltre 15 giorni dalla pubblicazione del presente avviso tramite posta elettronica certificata all’indirizzo PEC protocollo.ipsp@pec.cnr.it e per conoscenza a "),
                           ftext(RUP..E.mail), ftext(", indicando nell’oggetto “Att.ne "),
                           ftext(dott.rup), ftext(" "), ftext(RUP),
                           ftext(": preventivo relativo all’avviso pubblico per "),
                           ftext(la.fornitura), ftext(" di "), ftext(Prodotto), ftext("”.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("La documentazione trasmessa dovrà essere sottoscritta digitalmente con firma qualificata da un legale rappresentante/procuratore in grado di impegnare l’operatore economico.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Gli operatori economici stranieri non residenti in Italia, sprovvisti di posta elettronica certificata, possono inviare il preventivo e la dichiarazione in lingua italiana all’indirizzo "),
                         ftext(RUP..E.mail),
                         ftext(". Qualora l’O.E. straniero fosse sprovvisto di firma digitale dovrà sottoscrivere la dichiarazione con firma autografa e allegare alla dichiarazione un documento d’identità in corso di validità.")), style = "Normal") |>
        
        body_add_fpar(fpar(ftext("Individuazione dell'affidatario")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("L'individuazione dell'affidatario sarà operata discrezionalmente dalla Stazione Appaltante, nel caso in cui intenda procedere all’affidamento, a seguito dell'esame dei preventivi e delle relazioni tecniche ricevuti entro la scadenza.")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Non saranno presi in considerazione preventivi di importo superiore a quanto stimato dalla Stazione Appaltante.")), style = "Normal") |>
        
        body_add_fpar(fpar(ftext("Obblighi dell’affidatario")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("L’operatore economico affidatario sarà tenuto, prima dell’invio della lettera ordine, a fornire la seguente documentazione:")), style = "Normal")

      if(Importo.senza.IVA<40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Dichiarazione possesso requisiti di partecipazione e di qualificazione ai sensi del D.lgs. 36/2023;")), style = "Elenco punto")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("Dichiarazione possesso requisiti di qualificazione ai sensi del D.lgs. 36/2023;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("Comprova assolvimento imposta di bollo;")), style = "Elenco punto")
      }

      doc <- doc |>
        body_add_fpar(fpar(ftext("Patto di integrità ai sensi del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Comunicazione conto corrente dedicato ai sensi dell’art. 3, comma 7 della Legge 136/2010 e s.m.i.;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Dichiarazione di cui al DPCM 187/1991;")), style = "Elenco punto")
        
      if(CCNL!="Non applicabile"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Dettaglio del CCNL, stima degli oneri per la sicurezza dovuti a rischi da interferenze e dei costi aziendali della manodopera;")), style = "Elenco punto") |>
          body_add_fpar(fpar("Comprova dell’equivalenze delle tutele del CCNL utilizzato.", run_footnote(x=block_list(fpar(ftext(" Se l’OE applica un CCNL diverso da quello indicato dalla stazione appaltante.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Elenco punto")
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("[in caso di servizi e forniture per i quali è vigente un decreto sui CAM]", fpt.i),
                           ftext(" Documentazione attestante la conformità alle specifiche tecniche e alle clausole contrattuali contenute nei criteri ambientali minimi di cui al Decreto Ministeriale corrispondente")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("Eventuali procure.")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("L'operatore economico straniero non residente in Italia, invece, sarà tenuto a fornire solo la seguente documentazione:")), style = "Normal") |>
        body_add_fpar(fpar(ftext("Declaration on honour.")), style = "Elenco punto") |>
        
        body_add_fpar(fpar(ftext("Subappalto")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Fermi restando i limiti e le condizioni di ricorso al subappalto per le prestazioni secondarie ed accessorie, il subappalto delle prestazioni oggetto dell’affidamento, ai sensi dell’art. 119 co. 2 del Codice, può essere stipulato in misura non inferiore al 20 per cento delle prestazioni subappaltabili, con piccole e medie imprese, come definite dall’articolo 1, comma 1, lettera o) dell’allegato I.1. Gli operatori economici possono indicare una diversa soglia di affidamento delle prestazioni che si intende subappaltare alle piccole medie imprese per ragioni legate all’oggetto o alle caratteristiche delle prestazioni o al mercato di riferimento.")), style = "Normal") |>

        body_add_fpar(fpar(ftext("Chiarimenti")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("Per eventuali ri chieste inerenti il servizio e chiarimenti di natura procedurale/amministrativa l’operatore economico dovrà rivolgersi "),
                           ftext(al.rup),ftext(" referente della Stazione Appaltante, "),
                           ftext(dott.rup),ftext(" "),ftext(RUP),
                           ftext(", all’indirizzo e-mail "),ftext(RUP..E.mail),ftext(".")),style = "Normal") |>
        
        body_add_fpar(fpar(ftext("Trattamento dei dati personali")), style = "heading 3") |>
        body_add_fpar(fpar(ftext("I dati forniti dai soggetti proponenti come indicato nel documento allegato, saranno trattati ai sensi del Regolamento UE 679/2016 e, per quanto applicabile, ai sensi del D.lgs. 196/2003, come modificato dal D.lgs. 101/2018, esclusivamente per le finalità connesse all’espletamento del presente avviso.")), style = "Normal") |>
        body_add_par("", style = "Normal") |>
        body_add_fpar(fpar(ftext("Allegati:", fpt.b)), style = "Normal") |>
        body_add_fpar(fpar(ftext("1: Dichiarazione sostitutiva possesso requisiti OE per invio preventivo")), style = "Normal") |>
        body_add_fpar(fpar(ftext("2: Informativa sul trattamento dei dati personali")), style = "Normal") |>
        
        body_add_par("", style = "Normal") |>
        body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
        
      print(doc, target = "Avviso pubblico.docx")

      ## Allegato ----
      download.file(paste(lnk, "Vuoto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      doc <- doc |>
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
                           ftext("pienamente consapevole della responsabilità penale cui va incontro, ai sensi e per gli effetti dell’art. 76 D.P.R. 28 dicembre 2000, n. 445, in caso di dichiarazioni mendaci o di formazione, esibizione o uso di atti falsi ovvero di atti contenenti dati non più rispondenti a verità,")), style = "Normal") |>
        body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("di essere in possesso dei requisiti di cui all’avviso di indagine di mercato, e nello specifico:")), style = "Normal") |>
        body_add_fpar(fpar(ftext("abilitazione MePA relativa al Bando ___________, Categoria di abilitazione __________;"), 
                           run_footnote(x=block_list(fpar(ftext(" Riportare l’indicazione del bando di abilitazione utilizzato (esempio: Bando “Beni”, Bando “Servizi”) nonché la specifica Categoria merceologica. La categoria merceologica viene individuata attraverso la scelta del codice CPV.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript")),
                           ftext(";")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("requisiti di ordine generale di cui al Libro II, Titolo IV, Capo II del D.lgs. 36/2023;")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("requisiti d’idoneità professionale come specificato all’art. 100, comma 3 del D.lgs. n. 36/2023:")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("iscrizione nel registro della camera di commercio, industria, artigianato e agricoltura o nel registro delle commissioni provinciali per l’artigianato o presso i competenti ordini professionali per un’attività pertinente anche se non coincidente con l’oggetto dell’appalto. All’operatore economico di altro Stato membro non residente in Italia è richiesto di dichiarare ai sensi del testo unico delle disposizioni legislative e regolamentari in materia di documentazione amministrativa, di cui al decreto del Presidente della Repubblica del 28 dicembre 2000, n. 445 di essere iscritto in uno dei registri professionali o commerciali di cui all’allegato II.11 del D.lgs. 36/2023;")), style = "Elenco punto liv2") |>
        body_add_fpar(fpar(ftext("(eventuale)", fpt.i), ftext(" requisiti di capacità economico-finanziaria;")), style = "Elenco punto liv2") |>
        body_add_fpar(fpar(ftext("(eventuale)", fpt.i), ftext(" requisiti di capacità tecnico-organizzativa;")), style = "Elenco punto liv2") |>
        body_add_fpar(fpar(ftext("di essere iscritto in uno dei registri professionali e commerciali istituiti nel Paese in cui è residente;"), 
                           run_footnote(x=block_list(fpar(ftext(" Nel caso di operatori economici residenti in Paesi terzi firmatari dell'AAP o di altri accordi internazionali di cui all'art. 69 del D.Lgs 36/2023.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript")),
                           ftext(";")), style = "Elenco punto") |>
        body_add_fpar(fpar(ftext("documentate esperienze pregresse idonee all’esecuzione delle prestazioni contrattuali oggetto dell’affidamento.")), style = "Elenco punto") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext("Il sottoscritto dichiara, inoltre, di aver preso visione dell’informativa inerente il trattamento dei dati personali in conformità alla normativa vigente e in particolare al Regolamento GDPR 2016/679.")), style = "Normal") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext("Luogo e data ___________")), style = "Normal") |>
        body_add_par("") |>
        body_add_fpar(fpar("Firma digitale del legale rappresentante/procuratore", run_footnote(x=block_list(fpar(ftext("Per gli operatori economici italiani o stranieri residenti in Italia, la dichiarazione deve essere sottoscritta da un legale rappresentante ovvero da un procuratore del legale rappresentante, apponendo la firma digitale. Per gli operatori economici stranieri non residenti in Italia, la dichiarazione può essere sottoscritta dai medesimi soggetti apponendo la firma autografa ed allegando copia di un documento di identità del firmatario in corso di validità oppure con firma elettronica qualificata. Nel caso in cui la dichiarazione sia firmata da un procuratore del legale rappresentante, deve essere allegata copia conforme all’originale della procura oppure, nel solo caso in cui dalla visura camerale dell’operatore economico risulti l’indicazione espressa dei poteri rappresentativi conferiti con la procura, la dichiarazione sostitutiva resa dal procuratore/legale rappresentante sottoscrittore attestante la sussistenza dei poteri rappresentativi risultanti dalla visura.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
       
      print(doc, target = "Allegato 1 - Dichiarazione sostitutiva possesso requisiti OE per invio preventivo.docx")
      
      ## Privacy ----
      download.file(paste(lnk, "Privacy.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
      
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
      }
      
      doc <- doc |>
        cursor_bookmark("bookmark_oggetto") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext("La presente informativa descrive le misure di tutela riguardo al trattamento dei dati personali destinata ai fornitori di beni e/o servizi, nell’ambito dell’affidamento diretto "),
                           ftext(della.fornitura),
                           ftext(" di “"),
                           ftext(Prodotto, fpt.b),
                           ftext("”, ai sensi dell’articolo 13 del Regolamento UE 2016/679 in materia di protezione dei dati personali (di seguito, per brevità, GDPR).")), style = "Normal") |>
        body_replace_text_at_bkm("bookmark_oggetto_eng", Prodotto)
      print(doc, target = "Allegato 2 - Informativa privacy.docx")
    
      cat("

    Documenti generati:
          - Avviso pubblico
          - Allegato 1 - Dichiarazione sostitutiva possesso requisiti OE per invio preventivo
          - Allegato 2 - Informativa privacy
          ")

      ## Dati mancanti ---
      manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore)
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
  
  # RUP ----
  rup <- function(){
    if(fornitore.uscente=="vero"){
      cat(paste0(frase1, frase2, frase3.1, frase3.2, frase3.3, frase3.4, frase4))
      if(blocco.rota=="vero"){
        stop("Non è possibile continuare. Apportare le modifiche in FluOr come indicato sopra e, poi, generare nuovamente i documenti dopo aver scaricato Ordini.csv.\n")
      }else{
        cat("E' possibile continuare. Premere INVIO per proseguire\n")
        readline()
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
    #file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_reach("CAMPO.DELLA.FORNITURA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("OGGETTO", fpt.b),
                         ftext(": NOMINA DEL RESPONSABILE UNICO DEL PROGETTO AI SENSI DELL’ART. 15 E DELL’ALLEGATO I.2 DEL DECRETO LEGISLATIVO 31 MARZO 2023 N. 36 PER L’AFFIDAMENTO DIRETTO "),
                         ftext(toupper(della.fornitura)),
                         ftext(" DI “"),
                         ftext(toupper(Prodotto), fpt.b),
                         ftext("”"),
                         ftext(", ORDINE "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(", NELL'AMBITO DEL PROGETTO "),
                         ftext(Progetto.cup),
                         ftext(".")), style = "Normal") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      cursor_reach("CAMPO.NOMINE") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")
    
    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }
    
    doc <- doc |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      cursor_reach("CAMPO.RAS") |>
      body_remove()
    
    if(CCNL=="Non applicabile"){
      doc <- doc |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la "), ftext("richiesta di acquisto prot. ", fpt.b),
                         ftext(Prot..RAS, fpt.b), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”, nell’ambito del progetto "),
                         ftext(Progetto.cup),
                         ftext(", mediante affidamento diretto all'operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), per un importo presunto di "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA e di altre imposte e contributi di legge;")), style = "Normal")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la "), ftext("richiesta di acquisto prot. ", fpt.b),
                           ftext(Prot..RAS, fpt.b), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                           ftext(" relativa alla necessità di procedere all’acquisizione "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto),
                           ftext("”, nell’ambito del progetto "),
                           ftext(Progetto.cup),
                           ftext(", mediante affidamento diretto all'operatore economico "),
                           ftext(Fornitore),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), per un importo presunto di "),
                           ftext(Importo.senza.IVA),
                           ftext(", comprensivo di "),
                           ftext(Oneri.sicurezza),
                           ftext(" quali oneri per la sicurezza dovuti a rischi da interferenze, oltre IVA e di altre imposte e contributi di legge;")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATA", fpt.b), ftext(", pertanto, la necessità di procedere alla nomina del Responsabile Unico del Progetto per la progettazione, affidamento e esecuzione di una procedura di affidamento "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”;")), style = "Normal") |>
      cursor_bookmark("bookmark_rup") |>
      body_remove() |>
      cursor_backward()
    
    if(RUP!=RSS.nome){
      doc <- doc |>
      body_add_fpar(fpar(ftext("DI NOMINARE", fpt.b), ftext(" "), ftext(il.dott.rup), ftext(" "), ftext(RUP),
                         ftext(" Responsabile Unico del Progetto (RUP) che, ai sensi dell'art. 15 del Codice, dovrà:")), style = "Elenco liv1")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI ASSUMERE", fpt.b), ftext(" il ruolo di Responsabile Unico del Progetto (RUP) che, ai sensi dell'art. 15 del Codice, dovrà:")), style = "Elenco liv1")
    }
     
    doc <- doc |>
      body_add_fpar(fpar(ftext("svolgere tutte le attività indicate nell’allegato I.2 del Codice, o che siano comunque necessarie ove non di competenza di altri organi;")), style = "Elenco liv2")
      
    if(CCNL!="Non applicabile"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("individuare il CCNL, in base all’attività oggetto dell’appalto svolta dall’impresa, in conformità al comma 1 dell’art. 11 e allegato 1.01 del Codice, nonché ai sensi del comma 2-bis dell’art. 11 del Codice;")), style = "Elenco liv2")
    }
    
    if(Tipo.ordine=="Fuori MePA"){
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_pcp", "Piattaforma Contratti Pubblici (PCP)")
    }
    if(Supporto.RUP!=trattini){
      doc <- doc |>
        cursor_bookmark("bookmark_supporto_rup") |>
        body_remove() |>
        cursor_backward() |>
        body_add_fpar(fpar(ftext("DI INDIVIDUARE", fpt.b), ftext(" ai sensi dell’art. 15, comma 6 del Codice, "),
                           ftext(il.dott.sup), ftext(" "),
                           ftext(Supporto.RUP, fpt.b),
                           ftext(" in qualità di supporto al RUP, fermo restando i compiti e le mansioni a cui gli stessi sono già assegnati;")), style = "Elenco liv1")
    }else{
      doc <- doc |>
        cursor_bookmark("bookmark_supporto_rup") |>
        body_remove()
    }
      
    doc <- doc |>
      cursor_reach("CAMPO.FIRMA") |>
      body_remove() |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "2 Nomina RUP.docx"))
    
    cat("

    Documento generato: '2 Nomina RUP'")
    
    ## Dich. Ass. RSS ----
    download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    #file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50, COMMA 1, LETT. B) DEL D.LGS. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" DI “"),
                         ftext(Prodotto),
                         ftext("” nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", ordine "),
                         ftext(sede),
                         ftext(" "),
                         ftext(ordine),
                         ftext(y)), style = "Maiuscolo") |>
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
                         ftext(" all'operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext("), nell'ambito del progetto “"),
                         ftext(Progetto), ftext("”"), ftext(CUP1),
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
                         ftext(")")), style = "Firma 2")

    print(doc, target = paste0(pre.nome.file, "9.3 Dichiarazione assenza conflitto RSS.docx"))
    
    cat("

    Documento generato: '9.3 Dichiarazione assenza conflitto RSS'")
    
    ## Dich. Ass. RUP ----
    download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("All’"),
                         ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
      body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50, COMMA 1, LETT. B) DEL D.LGS. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" DI “"),
                         ftext(Prodotto),
                         ftext("”, nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", ordine "),
                         ftext(sede),
                         ftext(" "),
                         ftext(ordine),
                         ftext(y)), style = "Maiuscolo") |>
      body_add_fpar(fpar(ftext("DICHIARAZIONE DI ASSENZA DI SITUAZIONI DI CONFLITTO DI INTERESSI AI SENSI DEGLI ARTT. 46 e 47 D.P.R. 445/2000")), style = "heading 1") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar(ftext(sottoscritto.rup), ftext(" "), ftext(dott.rup), ftext(" "), ftext(RUP, fpt.b), ftext(", "),
                         ftext(nato.rup), ftext(" "), ftext(RUP..Luogo.di.nascita), ftext(", il "),
                         ftext(RUP..Data.di.nascita), ftext(", codice fiscale "), ftext(RUP..Codice.fiscale), ftext(", ")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la normativa attinente alle situazioni, anche potenziali, di conflitto di interessi, in relazione all’incarico di "),
                         ftext("Responsabile Unico del Progetto (RUP)", fpt.b),
                         ftext(" per l’affidamento "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ordine "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(", all’operatore economico "),
                         ftext(Fornitore, fpt.b),
                         ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext(")"),
                         ftext(", nell'ambito del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
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
      #body_add_fpar(fpar(ftext("")), style = "Normal") |>
      #body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da)), style = "Normal") |>
      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar("Il Responsabile Unico del Progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "9.4 Dichiarazione assenza conflitto RUP.docx"))
    cat("

    Documento generato: '9.4 Dichiarazione assenza conflitto RUP'")
    
    ## Dich. Ass. SUP ----
    if(Supporto.RUP!=trattini){
      download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
      
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
      }
      
      doc <- doc |>
        cursor_begin() |>
        body_add_fpar(fpar(ftext("All’"),
                           ftext("Istituto per la Protezione Sostenibile delle Piante", fpt.b)), style = "Destinatario", pos = "on") |>
        body_add_fpar(fpar(ftext("del Consiglio Nazionale delle Ricerche")), style = "Destinatario 2") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar(ftext("AFFIDAMENTO DIRETTO, AI SENSI DELL’ART. 50, COMMA 1, LETT. B) DEL D.LGS. N. 36/2023, "),
                           ftext(della.fornitura), ftext(" DI “"),
                           ftext(Prodotto),
                           ftext("” nell'ambito del progetto “"),
                           ftext(Progetto),
                           ftext("”"),
                           ftext(CUP1),
                           ftext(", ordine "),
                           ftext(sede),
                           ftext(" "),
                           ftext(ordine),
                           ftext(y)), style = "Maiuscolo") |>
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
                           ftext(" all'operatore economico "),
                           ftext(Fornitore, fpt.b),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext("), nell'ambito del progetto "),
                           ftext(Progetto.cup),
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
        body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da)), style = "Normal") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar("Il supporto al RUP", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.sup),
                           ftext(" "),
                           ftext(Supporto.RUP),
                           ftext(")")), style = "Firma 2")
        
      print(doc, target = paste0(pre.nome.file, "9.5 Dichiarazione assenza conflitto SUP.docx"))
      
      cat("

    Documento generato: '9.5 Dichiarazione assenza conflitto SUP'")
      
      # Patto integrità ----
      if(Fornitore..Nazione=="Italiana" & Fornitore..Rappresentante.legale!=trattini){
      download.file(paste(lnk, "Patto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      if(sede!="TOsi"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_rss", paste0(", che delega alla firma ", paste0(tolower(substr(firma.RSS, 1, 1)),substr(firma.RSS, 2, nchar(firma.RSS))),  " ", RSS))
      }
      
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_fornitura", paste0(della.fornitura, " di “", Prodotto, "” (", Pagina.web, "), nell'ambito del progetto ", Progetto1)) |>
        body_replace_text_at_bkm("bookmark_fornitore", paste0("L'operatore economico ", Fornitore, " (di seguito Operatore Economico) con sede legale in ", Fornitore..Sede, ", C.F./P.IVA ", as.character(Fornitore..P.IVA), ", rappresentato da ", Fornitore..Rappresentante.legale, " in qualità di ", tolower(Fornitore..Ruolo.rappresentante), ",")) |>
        body_replace_text_at_bkm("bookmark_firma", firma.RSS)
      
      print(doc, target = paste0(pre.nome.file, "5.1 Patto di integrità.docx"))
    }
    
      ## Dati mancanti ---
      manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore, RUP, Prot..RAS)
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
  
  
  # DaC ----
  dac <- function(){

    if(fornitore.uscente=="vero"){
      cat(paste0(frase1, frase2, frase3.1, frase3.2, frase3.3, frase3.4, frase4))
      if(blocco.rota=="vero"){
        stop("Non è possibile continuare. Apportare le modifiche in FluOr come indicato sopra e, poi, generare nuovamente i documenti dopo aver scaricato Ordini.csv.\n")
      }else{
        cat("E' possibile continuare. Premere INVIO per proseguire\n")
        readline()
      }
    }

    download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
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
                         ftext(", "),
                         ftext(Pagina.web),
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
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" lo Statuto del Consiglio Nazionale delle Ricerche, emanato con provvedimento del Presidente n. 93, prot. n. 0051080 del 19 luglio 2018, entrato in vigore in data 1° agosto 2018;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Regolamento di Organizzazione e Funzionamento del Consiglio Nazionale delle Ricerche emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 1° agosto 2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Regolamento di amministrazione contabilità e finanza, emanato con Provvedimento della Presidente n. 201 prot. n. 0507722 del 23 dicembre 2024, entrato in vigore dal 1° gennaio 2025 ed, in particolare, l’art.32 rubricato “Decisione di contrattare”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il D.lgs. 31 marzo 2023, n. 36 rubricato “Codice dei Contratti Pubblici in attuazione dell’articolo 1 della legge 21 giugno 2022, n. 78, recante delega al Governo in materia di contratti pubblici”, pubblicato sul Supplemento Ordinario n. 12 della GU n. 77 del 31 marzo 2023 (nel seguito per brevità “Codice”);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il D.lgs. 31 dicembre 2024, n. 209 rubricato “Disposizioni integrative e correttive al codice dei contratti pubblici, di cui al decreto legislativo 31 marzo 2023, n. 36”, pubblicato sul Supplemento Ordinario n.45/L della GU n. 305 del 31 dicembre 2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento n. 31 del Direttore Generale del CNR prot. n. 54747 del 25/2/2025 di attribuzione dell'incarico di Responsabile della Gestione e Compliance amministrativo-contabile (RGC) dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) alla sig.ra Concetta Mottura per il periodo dall’1/3/2025 al 29/2/2028;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall'art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 6 novembre 2012, n. 190 recante “Disposizioni per la prevenzione e la repressione della corruzione e dell’illegalità nella pubblica amministrazione” pubblicata sulla G.U.R.I. n. 265 del 13/11/2012;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il d.lgs. 14 marzo 2013, n. 33 recante “Riordino della disciplina riguardante il diritto di accesso civico e gli obblighi di pubblicità, trasparenza e diffusione di informazioni da parte delle pubbliche amministrazioni” pubblicato sulla Gazzetta Ufficiale n. 80 del 05/04/2013 e successive modifiche introdotte dal d.lgs. 25 maggio 2016 n. 97;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Codice di comportamento dei dipendenti del Consiglio Nazionale delle Ricerche approvato con delibera del Consiglio di Amministrazione n° 137/2017;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il vigente Piano triennale per la prevenzione della corruzione e della trasparenza (PTPCT) contenuto nel Piano Integrato di Attività e Organizzazione (PIAO), adottato con delibera del Consiglio di Amministrazione del Consiglio Nazionale delle Ricerche ai sensi dell’articolo 6 del decreto-legge n. 80/2021;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 23 dicembre 1999 n 488 e s.m.i., recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2000)”, ed in particolare l'articolo 26;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 27 dicembre 2006 n. 296, recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2007)”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la legge 24 dicembre 2007 n. 244 e s.m.i., recante “Disposizioni per la formazione del bilancio annuale e pluriennale dello Stato (Legge finanziaria 2008)”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il decreto-legge 7 maggio 2012 n. 52, convertito dalla legge 6 luglio 2012 n. 94 recante “Disposizioni urgenti per la razionalizzazione della spesa pubblica”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il decreto-legge 6 luglio 2012 n. 95, convertito con modificazioni dalla legge 7 agosto 2012 n. 135, recante “Disposizioni urgenti per la revisione della spesa pubblica con invarianza dei servizi ai cittadini”;")), style = "Normal")
      
    if(CCNL=="Non applicabile"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la richiesta di acquisto prot. "),
                           ftext(Prot..RAS), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                           ftext(" relativa alla necessità di procedere all’acquisizione "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto),
                           ftext("” ("),
                           ftext(Pagina.web),
                           ftext("), nell’ambito del progetto “"),
                           ftext(Progetto),
                           ftext("”"),
                           ftext(CUP1),
                           ftext(", mediante affidamento diretto all’operatore economico "),
                           ftext(Fornitore),
                           ftext(" (P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext(") per un importo presunto di "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA e di altre imposte e contributi di legge;")), style = "Normal")
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
                           ftext(") per un importo presunto di "),
                           ftext(Importo.senza.IVA),
                           ftext(", comprensivo di "),
                           ftext(Oneri.sicurezza),
                           ftext(" quali oneri per la sicurezza dovuti a rischi da interferenze ed "),
                           ftext(Manodopera),
                           ftext(" quali costi del personale, oltre IVA e di altre imposte e contributi di legge;")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento di nomina del Responsabile Unico del Progetto "),
                         ftext(dott.rup), ftext(" "), ftext(RUP),
                         ftext(", ai sensi dell’art. 15 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l'atto istruttorio prot. n. "),
                         ftext(Prot..atto.istruttorio), ftext(" con il quale il RUP ha dichiarato la conformità alle disposizioni di legge e ai regolamenti vigenti in materia del procedimento di selezione dell’affidatario proponendo, contestualmente, la formalizzazione dell’affidamento diretto in argomento all’OE "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(") per un importo pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" mediante atto immediatamente efficace;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che:")), style = "Normal") |>
      body_add_fpar(fpar(ftext("si intendono perseguire le finalità del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(" nell’ambito del quale è necessario acquisire la fornitura di cui trattasi, identificabile con il codice CPV "),
                         ftext(CPV),
                         ftext(";")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("alla data odierna non sono stati individuati, tra quelli messi a disposizione da CONSIP (Convenzioni, Accordi Quadro o Bandi del Sistema dinamico di acquisizione), strumenti idonei a soddisfare le già menzionate esigenze di approvvigionamento e che il bene/servizio oggetto della fornitura non è presente sulla Piattaforma regionale di riferimento;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("i "), ftext(beni), ftext(" di cui trattasi non sono presenti nel MePA;")), style = "Elenco punto")
    
    if(Tipo.ordine!="Fuori MePA" & Importo.senza.IVA.num>=5000){  
      doc <- doc |>
        body_add_fpar(fpar(ftext("trattandosi di beni funzionalmente destinati all’attività di ricerca d’importo pari o superiore a 5.000,00 euro trovano applicazioni le deroghe ai sensi dell’art. 4 comma 1 del D.L. 126/2019 convertito in L. 159/2019;")), style = "Elenco punto")
      }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("alla data odierna il "), ftext(bene), ftext(" oggetto della fornitura non è presente sulla Piattaforma regionale di riferimento;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("le prestazioni richieste non rientrano nell'elenco dei lavori, beni e servizi assoggettati a centralizzazione degli acquisti ai sensi dell'art.1 del Decreto del Presidente del Consiglio dei ministri del 16 agosto 2018;")), style = "Elenco punto")
    
    if(CCNL!="Non applicabile"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico ha confermato di applicare al personale dipendente impiegato nell’attività oggetto dell’appalto il contratto collettivo nazionale e territoriale di lavoro individuato dal RUP, ai sensi dell’art. 11, co. 2 del Codice, a seguito di autodichiarazione, individuato dai seguenti codice alfanumerico e codice ATECO "),
                           ftext(CCNL),
                           ftext(", stimando in "),
                           ftext(Manodopera),
                           ftext(" oltre IVA e altre imposte e contributi di legge i costi della manodopera"),
                           run_footnote(x=block_list(fpar(ftext(" Ai sensi dell’art. 2 dell’Allegato I.01 del Codice, è a cura del RUP l’individuazione del CCNL territoriale da applicare all’appalto secondo le procedure illustrate nell’Allegato I.01.", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript")),
                                        ftext(";")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che sono stati ritenuti congrui anche in considerazione della stima dei costi della manodopera effettuata dalla S.A., i costi della manodopera indicati dal già menzionato operatore economico a corredo dell’offerta, sulla base delle tariffe orarie previste per il CCNL codice alfanumerico e codice ATECO "),
                           ftext(CCNL),
                           ftext(";")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 50, comma 1, lettera b) del Codice, il quale consente, per affidamenti di contratti di servizi e forniture, ivi compresi i servizi di ingegneria e architettura e l'attività di progettazione di importo inferiore ad euro 140.000,00, di procedere ad affidamento diretto, anche senza consultazione di più operatori economici;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(", altresì, che la scelta dell’Operatore Economico deve essere effettuata assicurando che i soggetti individuati siano in possesso di documentate esperienze pregresse idonee all’esecuzione delle prestazioni contrattuali, anche individuati tra gli iscritti in elenchi o albi istituiti dalla stazione appaltante;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VALUTATA", fpt.b), ftext(" l’opportunità, in ottemperanza alla suddetta normativa, di procedere ad affidamento diretto all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(") mediante provvedimento contenente gli elementi essenziali descritti nell’art. 17, comma 2, del Codice, tenuto conto che il medesimo è in possesso di documentate esperienze pregresse idonee all’esecuzione della prestazione contrattuale;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO ", fpt.b),
                         ftext(rotazione.individuata)), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che dal 1° gennaio 2024 è fatto obbligo di utilizzare piattaforme di approvvigionamento digitale certificate (e-procurement) per svolgere le procedure di affidamento e di esecuzione dei contratti pubblici, a norma degli artt. 25 e 26 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che la stazione appaltante ai sensi dell’art. 48 comma 2 del Codice, ha accertato che il presente appalto non presenta un interesse transfrontaliero certo per cui non segue le procedure ordinarie di cui alla parte IV del Libro II;")), style = "Normal")

    if(Tipo.ordine=="Fuori MePA"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" il comunicato del Presidente dell’ANAC del 18 giugno 2025 di adozione del provvedimento di prolungamento della proroga di alcuni adempimenti previsti con la Delibera n. 582 del 13/12/2023 e con il Comunicato del Presidente del 18/12/2024, tra cui la  possibilità di utilizzare l’interfaccia web messa a disposizione dalla Piattaforma PCP dell’Autorità per gli affidamenti diretti di importo inferiore a 5.000 euro in caso di impossibilità o difficoltà di ricorso alle PAD al fine di consentire l’assolvimento delle funzioni ad essa demandate, ivi compresi gli obblighi in materia di trasparenza;")), style = "Normal")
    }
    # if(Motivo.fuori.MePA!="No"){
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(ICT.testo)), style = "Normal")
    # }

    doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che ai sensi dell’art. 53 comma 1 del Codice non sussistono particolari ragioni per la richiesta di garanzia provvisoria;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Bilancio Unico di Previsione del Consiglio Nazionale delle Ricerche per l’esercizio finanziario 2025, approvato dal Consiglio di Amministrazione con deliberazione n° 420/2024 del 17/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b), ftext(" la disponibilità finanziaria per la copertura della spesa sui fondi del progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", voce di costo CO.AN "),
                         ftext(Voce.di.spesa),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che vi sono i presupposti normativi e di fatto per acquisire "), ftext(la.fornitura), ftext(" in oggetto, nel rispetto dei principi generali enunciati nel Codice;")), style = "Normal") |>
      body_add_par("DISPONE", style = "heading 2")
    
    if(CCNL=="Non applicabile"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI PROCEDERE", fpt.b), ftext(" all’acquisizione "), ftext(della.fornitura), ftext(", mediante affidamento diretto ai sensi dell’art. 50, comma 1, lett. b) del Codice, all’operatore economico "),
                         ftext(Fornitore),
                         ftext(" con sede legale in "),
                         ftext(Fornitore..Sede),
                         ftext(", C.F. e P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", per un importo complessivo pari a euro "),
                         ftext(Importo.senza.IVA),
                         ftext(", oltre IVA e di altre imposte e contributi di legge;")), style = "Elenco liv1")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI PROCEDERE", fpt.b), ftext(" all’acquisizione "), ftext(della.fornitura), ftext(", mediante affidamento diretto ai sensi dell’art. 50, comma 1, lett. b) del Codice, all’operatore economico "),
                           ftext(Fornitore),
                           ftext(" con sede legale in "),
                           ftext(Fornitore..Sede),
                           ftext(", C.F. e P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext(", per un importo complessivo pari a euro "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA e di altre imposte e contributi di legge, comprensivo di "),
                           ftext(Oneri.sicurezza),
                           ftext(" quali oneri per la sicurezza dovuti a rischi da interferenze e "),
                           ftext(Manodopera),
                           ftext(" quali costi del personale;")), style = "Elenco liv1")
    }
      # body_add_fpar(fpar(ftext("DI NOMINARE ", fpt.b),
      #                    ftext(il.dott.rup),
      #                    ftext(" "),
      #                    ftext(RUP),
      #                    ftext(" Responsabile Unico del Progetto il quale, ai sensi dell’art. 15 del Codice, dovrà:")), style = "Elenco liv1") |>
    doc <- doc |>  
      body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" che l'affidamento di cui al presente provvedimento sia soggetto all’applicazione delle norme contenute nella legge n. 136/2010 e s.m.i. e che il pagamento venga disposto entro 30 giorni dall’emissione certificato di regolare esecuzione;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" in conformità a quanto disposto dall’art. 53, comma 4, del Codice, che l’affidatario non sarà tenuto a presentare la garanzia definitiva in quanto l'ammontare garantito sarebbe di importo così esiguo da non costituire reale garanzia per la stazione appaltante, determinando esclusivamente un appesantimento del procedimento;")), style = "Elenco liv1")
    
    if(CCNL!="Non applicabile"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" altresì che l’affidatario, ai sensi dell’art. 11 del Codice, sarà tenuto ad applicare il contratto collettivo nazionale e territoriale individuato dalla Stazione Appaltante e identificato dai seguenti codici alfanumerico e codice ATECO "),
                           ftext(CCNL),
                           ftext(";")), style = "Elenco liv1")
    }      
      
    doc <- doc |>
      body_add_fpar(fpar(ftext("DI SOTTOPORRE", fpt.b), ftext(" la lettera d’ordine alla condizione risolutiva in caso di accertamento della carenza dei requisiti di ordine generale;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("DI PROCEDERE", fpt.b), ftext(" con la registrazione sul sistema contabile della scrittura anticipata n. "),
                         ftext(Anticipata),
                         # ftext(" (da migrazione impegno in SIGLA n. "),
                         # ftext(N..impegno.di.spesa),
                         # ftext(")"),
                         ftext(" di "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA sul progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", voce di costo CO.AN "),
                         ftext(Voce.di.spesa),
                         ftext(", in favore dell'OE "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", soggetto registrato in U-Gov con il n. "),
                         ftext(Fornitore..Codice.terzo.SIGLA),
                         ftext(");")), style = "Elenco liv1") |>

    # if(Importo.senza.IVA.num>=40000){
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("DI STABILIRE", fpt.b), ftext(" che l'avvio dell'esecuzione del contratto o la sottoscrizione dello stesso/l’invio della lettera d’ordine siano subordinati all'esito della verifica dei requisiti di ordine generale, e speciale se previsti, senza rilevare cause ostative;")), style = "Elenco liv1") |>
    #     body_add_fpar(fpar(ftext("DI IMPEGNARE", fpt.b), ftext(" la spesa per un importo pari a € 35,00 sui fondi del già citato progetto, allocati sul GAE P___, voce COAN 13096 “Pubblicazione bandi di gara” per la contribuzione ANAC;")), style = "Elenco liv1")
    # }else{
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("DI SOTTOPORRE", fpt.b), ftext(" la lettera d’ordine alla condizione risolutiva in caso di accertamento della carenza dei requisiti di ordine generale;")), style = "Elenco liv1")
    # }

    #doc <- doc |>
      body_add_fpar(fpar(ftext("DI PROCEDERE", fpt.b), ftext(" alla pubblicazione del presente provvedimento ai sensi del combinato disposto dell’art. 37 del d.lgs. 14 marzo 2013, n. 33 e dell’art. 20 del Codice;")), style = "Elenco liv1") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("Controllo di regolarità contabile")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("Responsabile della Gestione e della Compliance amministrativo contabile (RGC)")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(Sig.ra Concetta Mottura)")), style = "Firma 1") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
      
    print(doc, target = paste0(pre.nome.file, "7 Decisione a contrattare.docx"))
    cat("

    Documento generato: '7 Decisione a contrattare'")

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore, RUP, Prot..RAS, Pagina.web, Prot..atto.istruttorio, Anticipata, CPV)
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

  # Provv. anticipata ----
  provv_ant <- function(){
    download.file(paste("https://raw.githubusercontent.com/giovabubi/appost/main/models/", "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("PROVVEDIMENTO")), pos = "on", style = "heading 1") |>
      body_add_fpar(fpar(ftext("Oggetto", fpt.b),
                         ftext(": registrazione sul sistema contabile di anticipata per il pagamento degli oneri derivanti dalla Richiesta di Acquisto "),
                         ftext(della.fornitura),
                         ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”, ordine "),
                         ftext(sede),
                         ftext(" "),
                         ftext(ordine),
                         ftext(y),
                         ftext(".")), style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo n. 127 del 4 giugno 2003, recante “Riordino del Consiglio Nazionale delle Ricerche”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo 31 dicembre 2009, n. 213 “Riordino degli Enti di Ricerca in attuazione dell’art. 1 della Legge 27 settembre 2007, n. 165”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo 25 novembre 2016, n. 218 “Semplificazione delle attività degli enti pubblici di ricerca ai sensi dell’articolo 13 della legge 7 agosto 2015, n. 124”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" la legge 7 agosto 1990, n. 241 recante “Nuove norme in materia di procedimento amministrativo e di diritto di accesso ai documenti amministrativi” pubblicata sulla Gazzetta Ufficiale n. 192 del 18/08/1990 e s.m.i.;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo 30 luglio 1999, n. 286 concernente “Riordino e potenziamento dei meccanismi e strumenti di monitoraggio e valutazione dei costi, dei rendimenti e dei risultati dell'attività svolta dalle amministrazioni pubbliche, a norma dell'articolo 11 della legge 15 marzo 1997, n. 59”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" lo Statuto del Consiglio Nazionale delle Ricerche, emanato con provvedimento del Presidente n. 93, prot. n. 0051080 del 19 luglio 2018, entrato in vigore in data 1° agosto 2018;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Organizzazione e Funzionamento del Consiglio Nazionale delle Ricerche - DPCNR n. 119 prot. n. 241776 del 10 luglio 2024, entrato in vigore dal 1° agosto 2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di amministrazione contabilità e finanza, emanato con Provvedimento della Presidente n. 201 prot. n. 0507722 del 23 dicembre 2024, entrato in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" la deliberazione del Consiglio di Amministrazione n. 201 del 28 giugno 2022 di approvazione del Piano di riorganizzazione e rilancio del Consiglio Nazionale delle Ricerche (CNR) che prevede il passaggio dalla contabilità finanziaria a quella economico-patrimoniale;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Provvedimento CNR n. 000114 del 30/10/2013 (prot. n. 0065484) relativo alla costituzione dell’Istituto per la Protezione Sostenibile delle Piante con successivi Provvedimento del Presidente n. 120 del 07/10/2014 (prot. n. 0072102) e Provvedimento. n. 26 del 29.03.22 di modifica e sostituzione del precedente atto costitutivo;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")
    
    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS2)), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Bilancio Unico di Previsione del Consiglio Nazionale delle Ricerche per l’esercizio finanziario 2025, approvato dal Consiglio di Amministrazione con deliberazione n. 420/2024 – Verbale 511 del 17/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la Richiesta di Acquisto prot. "),
                         ftext(Prot..RAS), ftext(" "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b),
                         ftext(" la disponibilità economico-finanziaria per la copertura degli oneri derivanti dall’acquisizione "),
                         ftext(della.fornitura),
                         ftext(" sui fondi del Progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce di costo COAN "),
                         ftext(Voce.di.spesa),
                         ftext(";")), style = "Normal") |>
      body_add_par("DISPONE", style = "heading 2") |>
      body_add_fpar(fpar(ftext("DI CONFERMARE", fpt.b),
                         ftext(" la registrazione sul sistema contabile dell'anticipata n. "),
                         ftext(Anticipata),
                         ftext(" di "),
                         ftext(Importo.con.IVA),
                         ftext(" IVA inclusa ("),
                         ftext(Importo.senza.IVA),
                         ftext(" IVA esclusa) sul progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce di costo COAN "),
                         ftext(Voce.di.spesa),
                         ftext(", in favore dell'OE "),
                         ftext(Fornitore),
                         ftext(" (P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(", soggetto U-Gov "),
                         ftext(Fornitore..Codice.terzo.SIGLA), 
                         ftext(").")), style = "Elenco punto") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(ftext("Controllo di regolarità amministrativa e contabile")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("Responsabile della Gestione e della Compliance amministrativo contabile (RGC)")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(Sig.ra Concetta Mottura)")), style = "Firma 1") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "3 Provv. anticipata.docx"))
    cat("

    Documento generato: '3 Provv. anticipata'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Importo.senza.IVA, Voce.di.spesa, Fornitore, Prot..RAS, Anticipata, Fornitore..Codice.terzo.SIGLA)
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
  
  # Provv. impegno ----
  provv_imp <- function(){
    download.file(paste("https://raw.githubusercontent.com/giovabubi/appost/main/models/", "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext(cdr, fpt.b)), style = "Normal") |>
      body_add_fpar(fpar(ftext("PROVVEDIMENTO DI ASSUNZIONE ANTICIPATA")), style = "heading 1") |>
      body_add_fpar(fpar(firma.RSS), style = "heading 2") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il d.lgs. 31 dicembre 2009 n. 213, recante “Riordino del Consiglio Nazionale delle Ricerche in attuazione dell’articolo 1 della Legge 27 settembre 2007, n. 165”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il d.lgs. 25 novembre 2016 n. 218, recante “Semplificazione delle attività degli enti pubblici di ricerca ai sensi dell'articolo 13 della legge 7 agosto 2015, n. 124”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Organizzazione e Funzionamento del CNR emanato con Provvedimento del Presidente nr. 119 Prot. n. 241776 del 10/07/2024, in vigore dal 01/08/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento CNR n. 114 del 30/10/2013 (prot. n. 0065484) relativo alla costituzione dell’Istituto per la Protezione Sostenibile delle Piante con successivi provvedimenti del Presidente n. 120 del 07/10/2014 (prot. n. 72102) e n. 2 del 11/01/2019 di conferma e sostituzione del precedente atto costitutivo;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }

    doc <- doc |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento n. 31 del Direttore Generale del CNR prot. n. 54747 del 25/2/2025 di attribuzione dell'incarico di Responsabile della Gestione e Compliance amministrativo-contabile (RGC) dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) alla sig.ra Concetta Mottura per il periodo dall’1/3/2025 al 29/2/2028;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la richiesta di acquisto prot. "),
                         ftext(Prot..RAS), ftext(" pervenuta "), ftext(dal.ric), ftext(" "), ftext(Richiedente),
                         ftext(" relativa alla necessità di procedere all’acquisizione "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”, nell’ambito delle attività previste dal progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         #ftext(", mediante affidamento diretto all’operatore economico "),
                         #ftext(Fornitore),
                         #ftext(" (P.IVA "),
                         #ftext(Fornitore..P.IVA),
                         #ftext(", soggetto U-Gov "),
                         #ftext(Fornitore..Codice.terzo.SIGLA),
                         #ftext(") per un importo stimato di "),
                         #ftext(Importo.senza.IVA),
                         #ftext(" oltre IVA;")), style = "Normal") |>
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la verifica effettuata della copertura finanziaria;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTA", fpt.b),
      #                    ftext(" la verifica del possesso da parte della Ditta aggiudicataria dei requisiti stabiliti dall’art. 94 D.Lgs. 36/2023;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b),
                         ftext(" che l'acquisizione in oggetto è funzionalmente destinata all’attività di ricerca;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("SENTITO", fpt.b),
                         ftext(" il parere del RUP che ha espletato un’adeguata indagine di mercato con la quale ha individuato la Ditta fornitrice alla quale affidare tramite affidamento diretto ai sensi dell’art. 50, comma 1, lett. b) del D.lgs. n. 36/2023;")), style = "Normal") |>
      body_add_par("DISPONE", style = "heading 2") |>
      body_add_fpar(fpar(ftext("l’affidamento "), ftext(della.fornitura), ftext(" alla Ditta che sarà aggiudicataria:")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("operatore economico: "), ftext(Fornitore), ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext(");")), style = "Elenco punto liv2") |>
      body_add_fpar(fpar(ftext("soggetto U-Gov: "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(";")), style = "Elenco punto liv2") |>
      #body_add_fpar(fpar(ftext("CUP: "), ftext(CUP2), ftext(";")), style = "Elenco punto liv2") |>
      body_add_fpar(fpar(ftext("l’assunzione della scrittura anticipata n. "),
                         ftext(Anticipata),
                         ftext(", di "),
                         ftext(Importo.con.IVA),
                         ftext(" IVA inclusa, con imputazione sulla voce COAN "),
                         ftext(Voce.di.spesa),
                         ftext(", progetto “"),
                         ftext(Progetto),
                         ftext("”"),
                         ftext(CUP1),
                         ftext(", natura istituzionale per "),
                         ftext(Importo.con.IVA),
                         ftext(" ;")), style = "Elenco punto")

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
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("di non inventariare il bene in quanto trattasi di materiale di consumo.")), style = "Elenco punto")
      }
    }else if(Tipo.acquisizione=='Servizi'){
      doc <- doc |>
        body_add_fpar(fpar(ftext("di non inventariare il bene in quanto trattasi di servizio.")), style = "Elenco punto")
    }

    doc <- doc |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")

    print(doc, target = paste0(pre.nome.file, "3 Provv. anticipata.docx"))
    cat("

    Documento generato: '3 Provv. anticipata'")

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore, Prot..RAS, Anticipata, Fornitore..Codice.terzo.SIGLA)
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

  # Richiesta pagina web ----
  pag <- function(){
    download.file(paste("https://raw.githubusercontent.com/giovabubi/appost/main/models/", "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
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
                         ftext(")")), style = "Firma 2")

    print(doc, target = paste0(pre.nome.file, "3 Richiesta pagina web.docx"))
    cat("

    Documento generato: '3 Richiesta pagina web'")

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

  # DocOE ----
  docoe <- function(){
    inpt.oe <- 1
#     if(ultimi.recente>0 & ultimi.recente<180){
#       cat(paste0("
# 
#       I documenti dell'operatore economico ", Fornitore, " sono già stati richiesti meno di 6 mesi fa (prot. ", ultimi.prot, ") in occasione dell'ordine n° ", ultimi.ordine, y,".
# Si vuole generare ugualmente i documenti dell'operatore economico per richiederli nuovamente?
#   1: Sì
#   2: No"))
#       inpt.oe <- readline()
#     }

    if(inpt.oe==1){
      if(Fornitore..Nazione=="Italiana"){

        ## Patto d'integrità ----
        download.file(paste(lnk, "Patto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
          doc <- read_docx("tmp.docx")
          file.remove("tmp.docx")
          
          if(sede!="TOsi"){
            doc <- doc |>
              body_replace_text_at_bkm("bookmark_rss", paste0(", che delega alla firma ", paste0(tolower(substr(firma.RSS, 1, 1)),substr(firma.RSS, 2, nchar(firma.RSS))),  " ", RSS))
          }
          
          doc <- doc |>
            body_replace_text_at_bkm("bookmark_fornitura", paste0(della.fornitura, " di “", Prodotto, "” (", Pagina.web, "), nell'ambito del progetto ", Progetto1)) |>
            body_replace_text_at_bkm("bookmark_fornitore", paste0("L'operatore economico ", Fornitore, " (di seguito Operatore Economico) con sede legale in ", Fornitore..Sede, ", C.F./P.IVA ", as.character(Fornitore..P.IVA), ", rappresentato da ", Fornitore..Rappresentante.legale, " in qualità di ", tolower(Fornitore..Ruolo.rappresentante), ",")) |>
            body_replace_text_at_bkm("bookmark_firma", firma.RSS)
          
          print(doc, target = paste0(pre.nome.file, "5.1 Patto di integrità.docx"))
        
        ## CC dedicato ----
        download.file(paste(lnk, "cc_dedicato.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        print(doc, target = paste0(pre.nome.file, "5.2 Comunicazione cc dedicato.docx"))

        ## DPCM ----
        download.file(paste(lnk, "DPCM.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_intestazione", int.docoe)
        print(doc, target = paste0(pre.nome.file, "5.3 Dichiarazione DPCM 187.docx"))
        
        cat("

    Documenti generati: '5.1 Patto di integrità', '5.2 Comunicazione cc dedicato', '5.3 Dichiarazione DPCM 187' e '5.4 Dichiarazione possesso requisiti di partecipazione e qualificazione'")

        ## Dati mancanti ---
        manca <- dplyr::select(sc, Fornitore, Prodotto, Progetto, Pagina.web)
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
        
        if(CCNL!="Non applicabile"){
          ## Manodopera ----
          download.file(paste(lnk, "Manodopera.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
          doc <- read_docx("tmp.docx")
          file.remove("tmp.docx")
          
          doc <- doc |>
            body_replace_text_at_bkm("bookmark_intestazione", int.docoe)
          print(doc, target = paste0(pre.nome.file, "5.5 Costi manodopera.docx"))
        }
        
          if(Importo.senza.IVA.num<40000){
            ## Part.Qual. ----
            download.file(paste(lnk, "Dich_requisiti_infra40.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
            doc <- read_docx("tmp.docx")
            file.remove("tmp.docx")
            
            doc <- doc |>
              body_replace_text_at_bkm("bookmark_intestazione", int.docoe)
            print(doc, target = paste0(pre.nome.file, "5.4 Dichiarazione possesso requisiti di partecipazione e qualificazione.docx"))
          }

            if(Importo.senza.IVA.num>=40000){
              ## Qual. ----
              download.file(paste(lnk, "Dich_requisiti_over40.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
              doc <- read_docx("tmp.docx")
              file.remove("tmp.docx")
              
              doc <- doc |>
                body_replace_text_at_bkm("bookmark_intestazione", int.docoe)
             
              print(doc, target = paste0(pre.nome.file, "5.4 Dichiarazione possesso requisiti di qualificazione.docx"))

              cat("
    Documento generato: '5.4 Dichiarazione possesso requisiti di qualificazione'")

              ## AUS ----
              download.file(paste(lnk, "Dich_aus.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
              doc <- read_docx("tmp.docx")
              file.remove("tmp.docx")
              
              doc <- doc |>
                body_replace_text_at_bkm("bookmark_intestazione", int.docoe)

              print(doc, target = paste0(pre.nome.file, "5.5 Dichiarazione del soggetto ausiliario.docx"))

              cat("
    Documento generato: '5.5 Dichiarazione del soggetto ausiliario'")
            }
        
        ## Condizioni d'acquisto ----
        download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
        
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
        }
        
        doc <- doc |>
          cursor_begin() |>
          body_add_par("CONDIZIONI GENERALI D'ACQUISTO", style = "heading 1", pos = "on") |>
          body_add_fpar(fpar(ftext("1. Ambito di applicazione", fpt.b), ftext(": le presenti condizioni generali di acquisto hanno la finalità di regolare in modo uniforme i rapporti contrattuali con i fornitori dai quali il CNR acquista beni e/o servizi in applicazione delle norme di legge e di regolamento. Le condizioni di vendita del fornitore non saranno in nessun caso applicabili ai rapporti contrattuali con il CNR, anche se fossero state richiamate in qualsiasi documento proveniente dal fornitore stesso.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("2. Resa", fpt.b), ftext(": franco destino.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("3. Durata", fpt.b), ftext(": "), ftext(fornitura.consegnata), ftext(" entro 30 giorni naturali e consecutivi decorrenti dalla data di sottoscrizione del presente contratto presso il luogo indicato nella pagina precedente.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("4. Subappalto", fpt.b), ftext(": in caso di subappalto trovano applicazione le disposizioni di cui all'art. 119 del codice dei contratti.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("5. Fatturazione", fpt.b), ftext(": la fattura, redatta secondo la normativa vigente, dovrà riportare, pena il rifiuto della stessa, il numero d'ordine (corrispondente al numero di registrazione al protocollo), il CIG e il CUP.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("6. Pagamento", fpt.b), ftext(": il pagamento sarà effettuato entro 30 gg. a partire dalla data del certificato di regolare esecuzione.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("7. Penali", fpt.b), ftext(": per ogni giorno naturale e consecutivo di ritardo rispetto ai termini previsti per l’esecuzione dell’appalto di cui all’art.8, si applicherà una penale pari all’1‰ (uno per mille) dell’importo contrattuale, al netto dell’IVA e dell’eventuale costo relativo alla sicurezza sui luoghi di lavoro derivante dai rischi di natura interferenziale. Per i soli contratti di forniture, nel caso in cui la prima verifica di conformità della fornitura abbia esito sfavorevole non si applicano le penali; qualora tuttavia l’Aggiudicatario non renda nuovamente la fornitura disponibile per la verifica di conformità entro i 20 (venti) giorni naturali e consecutivi successivi al primo esito sfavorevole, ovvero la verifica di conformità risulti nuovamente negativa, si applicherà la penale sopra richiamata per ogni giorno solare di ritardo. Nell’ipotesi in cui l’importo delle penali applicabili superi l’importo pari al 10% (dieci per cento) dell’importo contrattuale, al netto dell’IVA e dell’eventuale costo relativo alla sicurezza sui luoghi di lavoro derivante dai rischi di natura interferenziale, l’Ente risolverà il contratto in danno all’Aggiudicatario, salvo il diritto al risarcimento dell’eventuale ulteriore danno patito.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("8. Tracciabilità dei flussi finanziari", fpt.b), ftext(": il fornitore assume tutti gli obblighi di tracciabilità dei flussi finanziari di cui all’art. 3 della L. 136/2010 e s.m.i. Il mancato utilizzo del bonifico bancario o postale ovvero degli altri strumenti di incasso o pagamento idonei a consentire la piena tracciabilità delle operazioni costituisce motivo di risoluzione unilaterale del contratto. Il fornitore si impegna a consentire all’Amministrazione la verifica di cui al c. 9 art. 3 della legge 136/2010 e s.m.i. e a dare immediata comunicazione all'Amministrazione ed alla Prefettura-UTG della provincia ove ha sede l'Amministrazione della notizia dell’inadempimento della propria controparte (subappaltatore/subcontraente) agli obblighi di tracciabilità finanziaria.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("9. Osservanza leggi, regolamenti, contratti collettivi nazionali di lavoro, norme per la prevenzione infortuni ed igiene sul lavoro", fpt.b), ftext(": al personale impiegato nei servizi/forniture oggetto del presente appalto è applicato il contratto collettivo nazionale e territoriale in vigore per il settore e la zona nella quale si eseguono le prestazioni, stipulato dalle associazioni dei datori e dei prestatori di lavoro comparativamente più rappresentative sul piano nazionale e quello il cui ambito di applicazione sia strettamente connesso con l’attività oggetto dell’appalto svolta dall’impresa anche in maniera prevalente.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("L’obbligo permane anche dopo la scadenza dei suindicati contratti collettivi e fino alla loro sostituzione. L’affidatario che applica un differente contratto collettivo deve garantire le stesse tutele economiche e normative rispetto a quello indicato dalla Stazione Appaltante e come evidenziato nella dichiarazione di equivalenza presentata. I sopraccitati obblighi vincolano l’affidatario, anche nel caso che non sia aderente alle associazioni stipulanti o receda da esse ed indipendentemente dalla natura artigiana o industriale della struttura o dimensione della Società stessa e da ogni altra sua qualificazione giuridica, economica o sindacale, ivi compresa la forma cooperativa. L’affidatario è tenuto, inoltre, all’osservanza ed all’applicazione di tutte le norme relative alle assicurazioni obbligatorie ed antinfortunistiche, previdenziali ed assistenziali, nei confronti del proprio personale dipendente e dei soci lavoratori nel caso di cooperative. A richiesta della stazione appaltante, l’affidatario deve certificare l’applicazione del trattamento retributivo previsto dal CCNL delle imprese di settore e dagli accordi integrativi territoriali, ai lavoratori, compresi i soci lavoratori qualora si tratti di cooperativa, impiegati nell’appalto. La stazione appaltante si riserva di verificare, in qualsiasi momento, la regolarità dell’assolvimento degli obblighi inerenti al versamento dei contributi obbligatori ai sensi di legge. La stazione appaltante verifica, ai fini del pagamento della rata del corrispettivo, l’ottemperanza a tali obblighi, da parte dell’affidatario. La stazione appaltante si riserva di verificare, anche direttamente, il rispetto delle disposizioni in materia di assicurazioni obbligatorie per legge.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("Per inadempimenti contributivi o retributivi si applica il comma 6 dell’art. 11 del Codice.")), style = "Riquadro paragrafo")
        
        if(CCNL!="Non applicabile"){
          doc <- doc |>
            body_add_fpar(fpar(ftext("Il contratto collettivo nazionale e territoriale applicato è il seguente, identificato con codice alfanumerico e codice ATECO "),
                               ftext(CCNL),
                               ftext(".")), style = "Riquadro paragrafo")
        }
        
          doc <- doc |>
          body_add_fpar(fpar(ftext("10. Modifiche contrattuali", fpt.b), ftext(": la stazione appaltante, fermo quanto previsto dall’articolo sulla revisione dei prezzi [se presente], può modificare il contratto d’appalto conformemente a quanto disposto all'art.120 del codice dei contratti pubblici.")), style = "Riquadro paragrafo")
      
        if(Inventariabile=='Inventariabile'){
          doc <- doc |>
            body_add_fpar(fpar(ftext("11. Verifica di conformità", fpt.b), ftext(": la presente fornitura è soggetta a verifica di conformità da effettuarsi, secondo quanto previsto all'art. 116 e nell'allegato II.14 del codice dei contratti entro 1 mese. A seguito della verifica di conformità si procede al pagamento della rata di saldo e allo svincolo della cauzione.")), style = "Riquadro paragrafo")
        }
        if(Tipo.acquisizione=="Servizi"){
          doc <- doc |>
            body_add_fpar(fpar(ftext("11. Verifica di regolare esecuzione", fpt.b), ftext(": la stazione appaltante, per il tramite del RUP, emette il certificato di regolare esecuzione, secondo le modalità indicate nell'allegato II.14 al codice dei contratti pubblici, entro 1 mese. A seguito dell’emissione del certificato di regolare esecuzione si procede al pagamento della rata di saldo e allo svincolo della cauzione.")), style = "Riquadro paragrafo")
        }
        
        if(Importo.senza.IVA.num<40000){
          doc <- doc |>
            body_add_fpar(fpar(ftext("12. Clausola risolutiva espressa", fpt.b), ftext(": l’ordine è emesso in applicazione delle disposizioni contenute all’art. 52, commi 1 e 2 del d.lgs 36/2023. Il CNR ha diritto di risolvere il contratto/ordine in caso di accertamento della carenza dei requisiti di partecipazione. Per la risoluzione del contratto trovano applicazione l’art. 122 del d.lgs. 36/2023, nonché gli articoli 1453 e ss. del Codice Civile. Il CNR darà formale comunicazione della risoluzione al fornitore, con divieto di procedere al pagamento dei corrispettivi, se non nei limiti delle prestazioni già eseguite.")), style = "Riquadro paragrafo") |>
            body_add_fpar(fpar(ftext("13. Foro competente", fpt.b), ftext(": per ogni controversia sarà competente in via esclusiva il Tribunale di Roma.")), style = "Riquadro paragrafo")
        }else{
          doc <- doc |>
            body_add_fpar(fpar(ftext("12. Foro competente", fpt.b), ftext(": per ogni controversia sarà competente in via esclusiva il Tribunale di Roma.")), style = "Riquadro paragrafo")
        }
        
        doc <- doc |>
          body_add_par("") |>
          body_add_fpar(fpar(ftext("Le presenti condizioni generali di acquisto sono accettate mediante sovrascrizione, con firma digitale valida alla data di apposizione della stessa e a norma di legge.")), style = "Normal") |>
          body_add_par("") |>
          body_add_fpar(fpar("Per accettazione", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2")
        
        print(doc, target = paste0(pre.nome.file, "5.8 Condizioni acquisto.docx"))
        cat("
    Documento generato: '5.8 Condizioni acquisto'")
        
        ## Privacy ----
        download.file(paste(lnk, "Privacy.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
        
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
        }
        
        doc <- doc |>
          cursor_bookmark("bookmark_oggetto") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("La presente informativa descrive le misure di tutela riguardo al trattamento dei dati personali destinata ai fornitori di beni e/o servizi, nell’ambito dell’affidamento diretto "),
                             ftext(della.fornitura),
                             ftext(" di “"),
                             ftext(Prodotto, fpt.b),
                             ftext("”, ai sensi dell’articolo 13 del Regolamento UE 2016/679 in materia di protezione dei dati personali (di seguito, per brevità, GDPR).")), style = "Normal") |>
          body_replace_text_at_bkm("bookmark_oggetto_eng", Prodotto)
        print(doc, target = paste0(pre.nome.file, "5.9 Informativa privacy.docx"))
        cat("
    Documento generato: '5.9 Informativa privacy'")
        
      }else{
        ## Declaration on honour ----
        download.file(paste(lnk, "Honour.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        print(doc, target = paste0(pre.nome.file, "5.7 Declaration on honour.docx"))
        cat("

    Documento generato: '5.7 Declaration on honour'")
        
        ## Purchase conditions ----
        download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
        
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
        }
        
        doc <- doc |>
          body_add_par("GENERAL PURCHASE CONDITIONS", style = "heading 1", pos = "on") |>
          body_add_fpar(fpar(ftext("1. Scope of application", fpt.b), ftext(": These general conditions of purchase are intended to uniformly regulate contractual relationships with suppliers from whom CNR purchases goods and/or services in application of the laws and regulations. The supplier's conditions of sale will in no case be applicable to contractual relationships with CNR, even if they were referred to in any document originating from the supplier itself.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("2. Delivery", fpt.b), ftext(": to the destination.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("3. Duration", fpt.b), ftext(": the order must be delivered within 30 consecutive calendar days from the date of signing this contract at the location indicated on the previous page.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("4. Subcontracting", fpt.b), ftext(": in the event of subcontracting, the provisions of art. 119 of the Contracts Code apply.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("5. Invoice", fpt.b), ftext(": the invoice, drawn up in accordance with current legislation, must include, under penalty of rejection, the order number (corresponding to the protocol registration number), the CIG and the CUP.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("6. Payment", fpt.b), ftext(": payment will be made within 30 days from the date of the certificate of proper execution.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("7. Penalties", fpt.b), ftext(": for each natural and consecutive day of delay with respect to the terms provided for the execution of the contract referred to in art. 8, a penalty equal to 1‰ (one per thousand) of the contractual amount will be applied, net of VAT and any costs relating to safety in the workplace arising from risks of an interfering nature. For supply contracts only, in the event that the first conformity check of the supply has an unfavorable outcome, the penalties will not apply; however, if the Successful Bidder does not make the supply available again for the conformity check within 20 (twenty) natural and consecutive days following the first unfavorable outcome, or the conformity check is again negative, the penalty referred to above will be applied for each calendar day of delay. In the event that the amount of the applicable penalties exceeds the amount equal to 20% (twenty percent) of the contractual amount, net of VAT and any costs relating to safety in the workplace arising from interference risks, the Entity will terminate the contract to the detriment of the Successful Bidder, without prejudice to the right to compensation for any further damage suffered.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("8. Traceability of financial flows", fpt.b), ftext(": the supplier assumes all obligations of traceability of financial flows pursuant to art. 3 of Law 136/2010 and subsequent amendments. Failure to use bank or postal transfers or other collection or payment instruments suitable for allowing full traceability of transactions constitutes grounds for unilateral termination of the contract. The supplier undertakes to allow the Administration to carry out the verification pursuant to paragraph 9 of art. 3 of Law 136/2010 and subsequent amendments and to immediately notify the Administration and the Prefecture-UTG of the province where the Administration is based of the news of the failure of its counterpart (subcontractor/subcontractor) to comply with the obligations of financial traceability.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("9. Compliance with laws, regulations, national collective labor agreements, accident prevention and workplace hygiene standards", fpt.b), ftext(": the personnel employed in the services/supplies covered by this contract shall be subject to the national and territorial collective agreement in force for the sector and the area in which the services are performed, stipulated by the most representative employers' and workers' associations at national level and the one whose scope of application is strictly connected with the activity covered by the contract carried out by the company even in a prevalent manner.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("The obligation remains even after the expiry of the aforementioned collective agreements and until their replacement. The contractor who applies a different collective agreement must guarantee the same economic and regulatory protections compared to the one indicated by the Contracting Authority and as highlighted in the declaration of equivalence presented. The aforementioned obligations bind the contractor, even if it is not a member of the stipulating associations or withdraws from them and regardless of the artisan or industrial nature of the structure or size of the Company itself and of any other legal, economic or trade union qualification, including the cooperative form. The contractor is also required to comply with and apply all the rules relating to compulsory and accident prevention, social security and welfare insurance, with respect to its own employees and worker members in the case of cooperatives. At the request of the contracting authority, the contractor must certify the application of the remuneration treatment provided for by the CCNL of sector companies and by the territorial supplementary agreements, to the workers, including worker members in the case of a cooperative, employed in the contract. The contracting authority reserves the right to verify, at any time, the regularity of the fulfillment of the obligations relating to the payment of mandatory contributions pursuant to the law. The contracting authority verifies, for the purposes of payment of the instalment of the consideration, compliance with such obligations by the contractor. The contracting authority reserves the right to verify, even directly, compliance with the provisions regarding mandatory insurance by law.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("For non-compliance with contributions or wages, paragraph 6 of art. 11 of the Code applies.")), style = "Riquadro paragrafo") |>
          body_add_fpar(fpar(ftext("10. Contractual changes", fpt.b), ftext(": the contracting authority, without prejudice to the provisions of the article on price revision [if present], may modify the procurement contract in accordance with the provisions of art. 120 of the public procurement code.")), style = "Riquadro paragrafo")
        
        if(Inventariabile=='Inventariabile'){
          doc <- doc |>
            body_add_fpar(fpar(ftext("11. Compliance check", fpt.b), ftext(": this supply is subject to a conformity check to be carried out, as per art. 116 and Annex II.14 of the Contracts Code within 1 month. Following the conformity check, the balance instalment will be paid and the deposit will be released.")), style = "Riquadro paragrafo")
        }
        if(Tipo.acquisizione=="Servizi"){
          doc <- doc |>
            body_add_fpar(fpar(ftext("11. Verification of proper execution", fpt.b), ftext(": the contracting authority, through the RUP, issues the certificate of proper execution, according to the methods indicated in Annex II.14 to the Public Contracts Code, within 1 month. Following the issuance of the certificate of proper execution, the balance instalment is paid and the security is released.")), style = "Riquadro paragrafo")
        }
        
        if(Importo.senza.IVA.num<40000){
          doc <- doc |>
            body_add_fpar(fpar(ftext("12. Express termination clause", fpt.b), ftext(": the order is issued in application of the provisions contained in art. 52, paragraphs 1 and 2 of Legislative Decree 36/2023. The CNR has the right to terminate the contract/order in the event of a lack of participation requirements being ascertained. For the termination of the contract, art. 122 of Legislative Decree 36/2023, as well as articles 1453 et seq. of the Civil Code, apply. The CNR will formally communicate the termination to the supplier, with a ban on proceeding with the payment of the fees, except within the limits of the services already performed.")), style = "Riquadro paragrafo") |>
            body_add_fpar(fpar(ftext("13. Competent court", fpt.b), ftext(": the Court of Rome will have exclusive jurisdiction over any dispute.")), style = "Riquadro paragrafo")
        }else{
          doc <- doc |>
            body_add_fpar(fpar(ftext("12. Competent court", fpt.b), ftext(": the Court of Rome will have exclusive jurisdiction over any dispute.")), style = "Riquadro paragrafo")
        }
        
        doc <- doc |>
          body_add_par("") |>
          body_add_fpar(fpar(ftext("These general conditions of purchase are accepted by overwriting, with a digital signature valid on the date of affixing the same and in accordance with the law.")), style = "Normal") |>
          body_add_par("") |>
          body_add_fpar(fpar("Signature for acceptance", run_footnote(x=block_list(fpar(ftext(" The declarant must sign with a qualified digital signature or attach a photocopy of a valid identity document (art. 38 of Presidential Decree no. 445/2000 and subsequent amendments).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2")
        
        print(doc, target = paste0(pre.nome.file, "5.8 Purchase conditions.docx"))
        cat("
    Documento generato: '5.8 Purchase conditions'")
        
        ## Privacy eng ----
        download.file(paste(lnk, "Privacy_eng.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        file.remove("tmp.docx")
        
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
        
        if(sede=="TOsi"){
          doc <- doc |>
            headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
        }
        
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_oggetto_eng", Prodotto)
        
        print(doc, target = paste0(pre.nome.file, "5.9 Privacy policy.docx"))
        cat("
    Documento generato: '5.9 Privacy policy'")
      }
    }

        if(Importo.senza.IVA.num>=40000){
          ## Bollo ----
          download.file(paste(lnk, "Vuoto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
          doc <- read_docx("tmp.docx")
          file.remove("tmp.docx")
          
          doc <- doc |>
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
        Documento generato: '5.6 Comprova imposta di bollo'")
        }
    }
  
  # Comunicazione CIG ----
  com_cig <- function(){
    download.file(paste("https://raw.githubusercontent.com/giovabubi/appost/main/models/", "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
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
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2")
      
    print(doc, target = paste0(pre.nome.file, "6 Comunicazione CIG.docx"))
    cat("

    Documento generato: '6 Comunicazione CIG'")

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

  # AI ----
  ai <- function(){
    download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
    }
    
    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_fpar(fpar(ftext("ATTO ISTRUTTORIO")), style = "heading 1", pos = "on") |>
      body_add_fpar(fpar(ftext("Affidamento diretto, ai sensi dell’art. 50, comma 1, lett. b) del D.Lgs. N. 36/2023, "),
                         ftext(della.fornitura), ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”"),
                         ftext(CUI1),
                         ftext(", "), ftext(Pagina.web),
                         ftext(", nell'ambito del progetto “"),
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
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il D.lgs. 31 dicembre 2024, n. 209 rubricato “Disposizioni integrative e correttive al codice dei contratti pubblici, di cui al decreto legislativo 31 marzo 2023, n. 36”, pubblicato sul Supplemento Ordinario n.45/L della GU n. 305 del 31 dicembre 2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 50, comma 1, lettera b) del Codice, il quale consente, per affidamenti di contratti di servizi e forniture, ivi compresi i servizi di ingegneria e architettura e l'attività di progettazione di importo inferiore a euro 140.000,00, di procedere ad affidamento diretto, anche senza consultazione di più operatori economici;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento con il quale è stato nominato "),
                         ftext(il.dott.rup), ftext(" "), ftext(RUP),
                         ftext(" quale Responsabile Unico del Progetto ai sensi dell’art. 15 del Codice per l’affidamento di cui all’oggetto;")), style = "Normal")
    
    if(CCNL=="Non applicabile"){
      doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato "),
                         ftext(Fornitore),
                         ftext(" (C.F./P.IVA "),
                         ftext(Fornitore..P.IVA),
                         ftext(") ha presentato, attraverso la piattaforma telematica di negoziazione (RDO "),
                         ftext(as.character(RDO)),
                         ftext("), un’offerta ritenuta congrua corredata dalle dichiarazioni sostitutive richieste, in merito al possesso dei requisiti prescritti d’importo corrispondente al preventivo precedentemente acquisito e agli atti d’importo pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA e altre imposte e contributi di legge;")), style = "Normal")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato "),
                           ftext(Fornitore),
                           ftext(" (C.F./P.IVA "),
                           ftext(Fornitore..P.IVA),
                           ftext(") ha presentato, attraverso la piattaforma telematica di negoziazione (RDO "),
                           ftext(as.character(RDO)),
                           ftext("), un’offerta ritenuta congrua corredata dalle dichiarazioni sostitutive richieste, in merito al possesso dei requisiti prescritti d’importo corrispondente al preventivo precedentemente acquisito e agli atti d’importo pari a "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA e altre imposte e contributi di legge, comprensivo di "),
                           ftext(Oneri.sicurezza),
                           ftext(" quali oneri per la sicurezza dovuti a rischi da interferenze;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico "),
                           ftext(Fornitore),
                           ftext(" ha dichiarato che applicherà ai propri dipendenti il contratto collettivo nazionale e territoriale applicabile indentificato dai seguenti codice alfanumerico e codice ATECO "),
                           ftext(CCNL),
                           ftext(" indicato dalla Stazione Appaltante, ai sensi dell’art.11 del d.lgs.36/2023 e s.m.i.;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che i costi della manodopera indicati dal già menzionato operatore economico a corredo dell’offerta, sulla base delle tariffe orarie previste per il CCNL identificato dai codici alfanumerico e ATECO "),
                           ftext(CCNL),
                           ftext(" sono da ritenersi congrui anche in considerazione della stima dei costi della manodopera effettuata dalla S.A.;")), style = "Normal") |>
        body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" altresì i principi previsti dall’art. 57 del d.lgs. 36/2023 tra i quali le clausole sociali volte a garantire le pari opportunità generazionali, di genere e di inclusione lavorativa per le persone con disabilità o svantaggiate, la stabilità occupazionale del personale impiegato;")), style = "Normal")
    }

    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 52, comma 1 del Codice, il quale dispone che, nelle procedure di affidamento di cui all’art. 50, comma 1, lett. b) di importo inferiore a 40.000 euro, gli operatori economici attestano con dichiarazione sostitutiva di atto di notorietà il possesso dei requisiti di partecipazione e di qualificazione richiesti e che le stazioni appaltanti procedono alla risoluzione del contratto qualora a seguito delle verifiche non sia confermato il possesso dei requisiti generali dichiarati;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che, l’operatore economico individuato ha sottoscritto la dichiarazione sostitutiva di atto di notorietà resa ai sensi del D.P.R. n. 445/2000 attestante l’insussistenza di motivi di esclusione e il possesso dei requisiti di qualificazione richiesti;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che la Stazione appaltante, secondo il proprio regolamento interno, verificherà, previo sorteggio di un campione individuato con modalità predeterminata, le dichiarazioni degli operatori economici affidatari nelle procedure di affidamento di cui all’art. 50, comma 1, lett. b) di importo inferiore a 40.000 euro;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTI", fpt.b), ftext(" gli atti della procedura in argomento ed accertata la regolarità degli stessi in relazione alla normativa ed ai regolamenti vigenti;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VALUTATO", fpt.b), ftext(" il principio del risultato;")), style = "Normal") |>
      
    # if(Importo.senza.IVA.num<40000){
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" l’art. 52, comma 1 del Codice, il quale dispone che, nelle procedure di affidamento di cui all’art. 50, comma 1, lett. b) di importo inferiore a 40.000 euro, gli operatori economici attestano con dichiarazione sostitutiva di atto di notorietà il possesso dei requisiti di partecipazione e di qualificazione richiesti e che le stazioni appaltanti procedono alla risoluzione del contratto qualora a seguito delle verifiche non sia confermato il possesso dei requisiti generali dichiarati;")), style = "Normal") |>
    #     body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato ha sottoscritto la dichiarazione sostitutiva attestante il possesso dei requisiti di ordine generale previsti dal Codice ai sensi dell’art. 52 del Codice, archiviata con prot. ")), style = "Normal") |>
    #     #                   ftext(Prot..DocOE), ftext(";")), style = "Normal") |>
    #     body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che la Stazione appaltante verificherà, previo sorteggio di un campione individuato con modalità predeterminata, le dichiarazioni degli operatori economici affidatari;")), style = "Normal")
    # }else{
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato ha sottoscritto la dichiarazione sostitutiva attestante il possesso dei requisiti di ordine generale previsti dal Codice ai sensi dell’art. 52 del Codice e il DGUE ai fini dell’avvio delle verifiche ai sensi dell’art. 94, 95, 96, 97, 98 e 100 del d.lgs. n. 36/2023 e successive modifiche ed integrazioni;")), style = "Normal") |>
    #     body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che le verifiche effettuate ai sensi dell’art. 94, 95, 96, 97, 98 e 100 del d.lgs. n. 36/2023 non hanno rilevato cause ostative nei confronti dell’operatore economico individuato;")), style = "Normal")
    # }

      body_add_fpar(fpar(ftext("AI FINI DELL’ISTRUTTORIA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("Dichiara che il procedimento di selezione dell’affidatario risulta condotto in conformità alle disposizioni di legge e ai regolamenti vigenti in materia;")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("Propone il perfezionamento dell’affidamento diretto nei confronti dell’operatore economico "),
                         ftext(Fornitore), ftext(" (C.F./P.IVA "), ftext(Fornitore..P.IVA),
                         ftext(") per un importo complessivo pari a "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA mediante provvedimento di decisione di contrattare immediatamente efficace.")), style = "Elenco liv1") |>
      
    # if(Importo.senza.IVA.num<40000){
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("Nulla osta all’emissione della lettera d’ordine purché munita di apposita clausola risolutiva in caso di accertamento della carenza dei requisiti di ordine generale.")), style = "Elenco liv1")
    # }else{
    #   doc <- doc |>
    #     body_add_fpar(fpar(ftext("Nulla osta al perfezionamento della lettera d’ordine/contratto con l’Operatore Economico individuato.")), style = "Elenco liv1")
    # }

      body_add_fpar(fpar(ftext("")), style = "Normal") |>
      body_add_fpar(fpar("Il Responsabile Unico del Progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("),
                         ftext(dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "5 Atto istruttorio.docx"))
    cat("

    Documento generato: '5 Atto istruttorio'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Fornitore, RUP, Prot..RAS, Pagina.web, Fornitore..Codice.terzo.SIGLA, RDO)
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

  # Lettera d'ordine ----
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

    download.file(paste(lnk, "LdO.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1) |>
      headers_replace_text_at_bkm("bookmark_headers_sede_en", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale") |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale_en", "Istituzionale")
    }

    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("LETTERA D’ORDINE "), ftext(sede), ftext(" "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
      body_replace_text_at_bkm("bookmark_cup", CUP2) |>
      body_replace_text_at_bkm("bookmark_cig", CIG) |>
      body_replace_text_at_bkm("bookmark_cui", CUI2) |>
      body_replace_text_at_bkm("bookmark_rup", RUP) |>
      body_replace_text_at_bkm("bookmark_offerta", Preventivo.fornitore) |>
      body_replace_text_at_bkm("bookmark_dac", Prot..DaC) |>
      body_replace_text_at_bkm("bookmark_rdo1", ordine.trattativa.scelta.ldo1) |>
      body_replace_text_at_bkm("bookmark_rdo2", as.character(ordine.trattativa.scelta.ldo2)) |>
      body_replace_text_at_bkm("bookmark_web", Pagina.web) |>
      body_replace_text_at_bkm("bookmark_progetto", Progetto) |>
      body_replace_text_at_bkm("bookmark_fornitore", Fornitore) |>
      body_replace_text_at_bkm("bookmark_sede", Fornitore..Sede) |>
      body_replace_text_at_bkm("bookmark_piva", as.character(Fornitore..P.IVA)) |>
      body_replace_text_at_bkm("bookmark_pec", Fornitore..PEC) |>
      body_replace_text_at_bkm("bookmark_email", Fornitore..E.mail) |>
      cursor_bookmark("bookmark_tabella") |>
      body_add_table(prt, style = "Tabella LdO", pos = "on") |>
      body_replace_text_at_bkm("bookmark_imponibile", Imponibile.ldo.txt) |>
      body_replace_text_at_bkm("bookmark_aliquota", paste0("IVA (", Aliquota.IVA, ")")) |>
      body_replace_text_at_bkm("bookmark_iva", IVA.ldo.txt) |>
      body_replace_text_at_bkm("bookmark_importo", Importo.ldo.txt) |>
      body_replace_text_at_bkm("bookmark_consegna", Richiedente..Luogo.di.consegna) |>
      body_replace_text_at_bkm("bookmark_cuu", CUU) |>
      body_replace_text_at_bkm("bookmark_fatturazione", fatturazione) |>
      body_replace_text_at_bkm("bookmark_note", paste0("ordine n° ", sede, " ", ordine, y, ", prot. n. _____ (si veda in alto nella pagina della lettera d'ordine), CIG ", CIG, ", CUP ", CUP, ", progetto ", Progetto, ".")) |>
      cursor_bookmark("bookmark_firma") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2", pos = "on") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
      
    if(Fornitore..Nazione=="Italiana"){
      b <- cursor_bookmark(doc, "bookmark_ldo_en")
      b <- doc$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which
      doc <- cursor_bookmark(doc, "bookmark_ldo_en")
      for(i in 1:(e-b)){
        doc <- body_remove(doc)
      }
    }else{
      doc <- doc |>
        cursor_bookmark("bookmark_ldo_en") |>
        body_add_fpar(fpar(ftext("PURCHASE ORDER "), ftext(sede), ftext(" N° "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
        body_replace_text_at_bkm("bookmark_en_cup", CUP2) |>
        body_replace_text_at_bkm("bookmark_en_cig", CIG) |>
        body_replace_text_at_bkm("bookmark_en_cui", CUI2) |>
        body_replace_text_at_bkm("bookmark_en_rup", RUP) |>
        body_replace_text_at_bkm("bookmark_en_offerta", Preventivo.fornitore) |>
        body_replace_text_at_bkm("bookmark_en_dac", Prot..DaC.en) |>
        body_replace_text_at_bkm("bookmark_en_rdo1", ordine.trattativa.scelta.ldo1) |>
        body_replace_text_at_bkm("bookmark_en_rdo2", as.character(ordine.trattativa.scelta.ldo2)) |>
        body_replace_text_at_bkm("bookmark_en_web", Pagina.web) |>
        body_replace_text_at_bkm("bookmark_en_progetto", Progetto) |>
        body_replace_text_at_bkm("bookmark_en_fornitore", Fornitore) |>
        body_replace_text_at_bkm("bookmark_en_sede", Fornitore..Sede) |>
        body_replace_text_at_bkm("bookmark_en_piva", as.character(Fornitore..P.IVA)) |>
        body_replace_text_at_bkm("bookmark_en_pec", Fornitore..PEC) |>
        body_replace_text_at_bkm("bookmark_en_email", Fornitore..E.mail) |>
        cursor_bookmark("bookmark_en_tabella") |>
        body_add_table(prt.en, style = "Tabella LdO", pos = "on") |>
        body_replace_text_at_bkm("bookmark_en_imponibile", Importo.senza.IVA) |>
        body_replace_text_at_bkm("bookmark_en_aliquota", paste0("VAT (", Aliquota.IVA, ")")) |>
        body_replace_text_at_bkm("bookmark_en_iva", IVA) |>
        body_replace_text_at_bkm("bookmark_en_importo", Importo.con.IVA) |>
        body_replace_text_at_bkm("bookmark_en_consegna", Richiedente..Luogo.di.consegna) |>
        body_replace_text_at_bkm("bookmark_cuu_en", CUU) |>
        body_replace_text_at_bkm("bookmark_en_fatturazione", fatturazione) |>
        body_replace_text_at_bkm("bookmark_en_note", paste0("purchase order no. ", sede, " ", ordine, y, ", prot. n. _____ (see on the top of this page), CIG ", CIG, ", CUP ", CUP, ", project ", Progetto, ".")) |>
        cursor_bookmark("bookmark_en_firma") |>
        body_add_fpar(fpar(ftext("The Responsible")), style = "Firma 2", pos = "on") |>
        body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    }
    print(doc, target = paste0(pre.nome.file, "8 Lettera ordine.docx"))

    cat("

    Documento generato: '8 Lettera ordine'")

    ## Dati mancanti ---
    manca <- dplyr::select(sc, CIG, RUP, RDO, Fornitore, Importo.senza.IVA, Aliquota.IVA, Pagina.web, Prot..DaC)
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

  # Dich. Prestazione resa ----
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

    
    download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale", only_at_cursor = TRUE)
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
                         ftext(Fornitore), ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext("; soggetto U-Gov "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(");")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il documento di trasporto "),
                         ftext(DDT),
                         ftext(";")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("DICHIARA")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("di aver svolto la procedura secondo la normativa vigente;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext(materiale.conforme)), style = "Elenco punto") |>
      # body_add_par("") |>
      # body_add_fpar(fpar(ftext(sede1), ftext(", __/__"), ftext(y)), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Il Responsabile Unico del Progetto (RUP)")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(dott.rup), ftext(" "), ftext(RUP), ftext(")")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "10 Dichiarazione prestazione resa.docx"))
    cat("

    Documento generato: '10 Dichiarazione prestazione resa'")

    ## Dati mancanti ---
    manca <- dplyr::select(sc, CIG, RUP, RDO, Fornitore, Importo.senza.IVA, Aliquota.IVA, Pagina.web, Prot..DaC, DDT, Fornitore..Codice.terzo.SIGLA)
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

  # Regolare esecuzione ----
  reg_es <- function(){
    if(PNRR!="No"){
      download.file(paste(lnk, "Vuoto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove("tmp.docx")
      file.remove(logo)
    }else{
      download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
      }
    }
    
    doc <- doc |>
      cursor_begin() |>
      cursor_forward() |>
      body_add_par("CERTIFICATO DI REGOLARE ESECUZIONE DELLE PRESTAZIONI", style = "heading 1", pos = "on") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("OGGETTO: ", fpt.b), ftext(bene), ftext(" di “"), ftext(Prodotto), ftext("”.")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CIG ", fpt.b), ftext(CIG)), style = "Normal") |>
      body_add_fpar(fpar(ftext("CUP ", fpt.b), ftext(CUP)), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Importo: ", fpt.b), ftext(Importo.senza.IVA), ftext(" oltre IVA")), style = "Normal") |>
      body_add_fpar(fpar(ftext("RUP: ", fpt.b), ftext(RUP)), style = "Normal") |>
      # body_add_fpar(fpar(ftext("Soggetto: "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(" - "),
      #                    ftext(Fornitore),
      #                    ftext(" - P.IVA "),
      #                    ftext(Fornitore..P.IVA)), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(sottoscritto.rup), ftext(", in qualità di Responsabile Unico del Progetto nominato con provvedimento del "), ftext(sub("...","",firma.RSS)), 
                         ftext(" prot. n. "), ftext(Prot..nomina.RUP), ftext(" per "), ftext(la.fornitura), ftext(" in oggetto,")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la RDO MePA n. "), ftext(as.character(RDO)),
                         ftext(" relativa all'acquisizione "), ftext(della.fornitura), ftext(" di “"), ftext(Prodotto), ftext("”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la lettera d'ordine "), ftext(sede), ftext(" "), ftext(ordine), ftext(y),
                         ftext(" prot. n. "), ftext(Prot..lettera.ordine), ftext(";")), style = "Normal")
    if(Tipo.acquisizione=="Beni"){
      doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il DDT "), ftext(DDT), ftext(" di consegna parziale/totale e relativa fattura;")), style = "Normal")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la fattura _____ relativa al servizio di “"), ftext(Prodotto), ftext("”;")), style = "Normal")
    }
    doc <- doc |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che "), ftext(la.fornitura), ftext(" in argomento "), ftext(fornitura.eseguita), ftext(" entro i termini indicati nella lettera d'ordine;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" che il richiedente l'acquisto "), ftext(dott.ric), ftext(" "), ftext(Richiedente),
                         ftext(" ha verificato la conformità ed il perfetto funzionamento di quanto richiesto;")), style = "Normal") |>
      body_add_par("CERTIFICA", style = "heading 2") |>
      body_add_fpar(fpar(ftext("l’esecuzione a regola d’arte delle prestazioni affidate sotto il profilo tecnico e funzionale, in conformità e nel rispetto delle condizioni, modalità, termini e prescrizioni della lettera d’ordine, nonché nel rispetto delle disposizioni vigenti in materia;")), style = "Elenco punto")
    if(Tipo.acquisizione=="Beni"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("la conformità al tipo, ai modelli e alle quantità descritte nell’offerta e la rispondenza alle caratteristiche tecniche, economiche e quantitative nel rispetto delle previsioni dedotte nei documenti di affidamento;")), style = "Elenco punto")
    }
    doc <- doc |>
      body_add_fpar(fpar(ftext("la regolare esecuzione della prestazione nonché il rilascio, da parte dell’esecutore, della completa documentazione e di quanto espressamente richiesto come elemento "),
                         ftext(della.fornitura), ftext(" in oggetto;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("il rilascio di tutte le certificazioni richieste e delle garanzie previste dalle vigenti disposizioni;")), style = "Elenco punto") |>
      body_add_fpar(fpar(ftext("il presente certificato si rilascia in ottemperanza alla normativa vigente ai fini delle consequenziali attività amministrative e contabili.")), style = "Elenco punto") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Il Responsabile Unico del Progetto (RUP)")), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(dott.rup), ftext(" "), ftext(RUP), ftext(")")), style = "Firma 2")
      
      print(doc, target = paste0(pre.nome.file, "8 Certificato regolare esecuzione.docx"))
      cat("

    Documento generato: '8 Certificato regolare esecuzione'")
      
      ## Dati mancanti ---
      manca <- dplyr::select(sc, Prodotto, Progetto, Fornitore, CIG, CUP, Voce.di.spesa, DDT, Prot..nomina.RUP, Prot..lettera.ordine, RDO)
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
  
  
  # Provv. Liquidazione ----
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

    download.file(paste(lnk, "Intestata.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    if(PNRR!="No"){
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      file.remove(logo)
    }else{
      doc <- doc |>
        headers_replace_text_at_bkm("bookmark_headers_sede", sede1)
      if(sede=="TOsi"){
        doc <- doc |>
          headers_replace_text_at_bkm("bookmark_headers_istituzionale", "Istituzionale")
      }
    }

    doc <- doc |>
      body_add_par("PROVVEDIMENTO DI LIQUIDAZIONE E PAGAMENTO", style = "heading 1", pos = "on") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("OGGETTO:", fpt.b)), style = "Normal") |>
      body_add_fpar(fpar(ftext("Fattura n. _____ del _____ - Importo totale del documento "), ftext(Importo.con.IVA)), style = "Normal") |>
      body_add_fpar(fpar(ftext("Progetto: "), ftext(Progetto.int.no.cup)), style = "Normal") |>
      body_add_fpar(fpar(ftext("CIG "), ftext(CIG), ftext(" - CUP "), ftext(CUP)), style = "Normal") |>
      body_add_fpar(fpar(ftext("Soggetto: " ,fpt.b), ftext(Fornitore..Codice.terzo.SIGLA, fpt.b), ftext(" - ", fpt.b),
                         ftext(Fornitore, fpt.b),
                         ftext(" - P.IVA/C.F. ", fpt.b),
                         ftext(Fornitore..P.IVA, fpt.b)), style = "Normal") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "heading 1") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo n. 127 del 4 giugno 2003, recante “Riordino del Consiglio Nazionale delle Ricerche”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo 31 dicembre 2009, n. 213 “Riordino degli Enti di Ricerca in attuazione dell’art. 1 della Legge 27 settembre 2007, n. 165”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo 25 novembre 2016, n. 218 “Semplificazione delle attività degli enti pubblici di ricerca ai sensi dell’articolo 13 della legge 7 agosto 2015, n. 124”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" la legge 7 agosto 1990, n. 241 recante “Nuove norme in materia di procedimento amministrativo e di diritto di accesso ai documenti amministrativi” pubblicata sulla Gazzetta Ufficiale n. 192 del 18/08/1990 e s.m.i.;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Decreto Legislativo 30 luglio 1999, n. 286 concernente “Riordino e potenziamento dei meccanismi e strumenti di monitoraggio e valutazione dei costi, dei rendimenti e dei risultati dell'attività svolta dalle amministrazioni pubbliche, a norma dell'articolo 11 della legge 15 marzo 1997, n. 59”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" lo Statuto del Consiglio Nazionale delle Ricerche, emanato con provvedimento del Presidente n. 93, prot. n. 0051080 del 19 luglio 2018, entrato in vigore in data 1° agosto 2018;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di Organizzazione e Funzionamento del Consiglio Nazionale delle Ricerche - DPCNR n. 119 prot. n. 241776 del 10 luglio 2024, entrato in vigore dal 1° agosto 2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Regolamento di amministrazione contabilità e finanza, emanato con Provvedimento della Presidente n. 201 prot. n. 0507722 del 23 dicembre 2024, entrato in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" la deliberazione del Consiglio di Amministrazione n. 201 del 28 giugno 2022 di approvazione del Piano di riorganizzazione e rilancio del Consiglio Nazionale delle Ricerche (CNR) che prevede il passaggio dalla contabilità finanziaria a quella economico-patrimoniale;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Provvedimento CNR n. 000114 del 30/10/2013 (prot. n. 0065484) relativo alla costituzione dell’Istituto per la Protezione Sostenibile delle Piante con successivi Provvedimento del Presidente n. 120 del 07/10/2014 (prot. n. 0072102) e Provvedimento. n. 26 del 29/03/2022 di modifica e sostituzione del precedente atto costitutivo;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")

    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }
    
    doc <- doc |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), 
                         #ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, UO 121.000) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
                         #ftext(" il provvedimento n. 31 del Direttore Generale del CNR prot. n. 54747 del 25/2/2025 di attribuzione dell'incarico di Responsabile della Gestione e Compliance amministrativo-contabile (RGC) dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) alla sig.ra Concetta Mottura per il periodo dall’1/3/2025 al 29/2/2028;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b),
      #                    ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b),
                         ftext(" il Bilancio Unico di Previsione del Consiglio Nazionale delle Ricerche per l’esercizio finanziario 2025, approvato dal Consiglio di Amministrazione con deliberazione n° 420/2024 – Verbale 511 del 17/12/2024;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la richiesta d'acquisto prot. n. "),
                         ftext(Prot..RAS),
                         ftext(", registrazione U-GOV anticipata n. "),
                         ftext(Anticipata),
                         ftext(", voce di costo CO.AN "),
                         ftext(Voce.di.spesa),
                         ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" il provvedimento di decisione a contrattare prot. n. "),
                         ftext(Prot..DaC),
                         ftext(" con il quale è stato disposto l'acquisto "),
                         ftext(della.fornitura),
                         ftext(" di “"),
                         ftext(Prodotto),
                         ftext("”;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTA", fpt.b),
                         ftext(" la lettera d’ordine "), ftext(sede),
                         ftext(" "), ftext(ordine), ftext(y),
                         ftext(", prot. n. "),
                         ftext(Prot..lettera.ordine),
                         ftext(", RDO MePA "),
                         ftext(as.character(RDO)),
                         ftext(", CIG "),
                         ftext(CIG),
                         ftext(", registrazione U-Gov anticipata n. __________;")), style = "Normal")
    if(Fornitore..Nazione=="Italiana"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b),
                           ftext(" la fattura elettronica identificativo SDI n. _____, registrazione U-GOV scrittura normale n. _____ del _____;")), style = "Normal")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b),
                           ftext(" la fattura estera cartacea prot. n. _____, registrazione U-GOV scrittura normale n. _____ del _____;")), style = "Normal")
    }
      
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il DDT "),
                                                    ftext(DDT),
                                                    ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("ACCERTATO", fpt.b), ftext(" il riscontro positivo sulla regolarità dell’esecuzione della prestazione e sulla rispondenza della stessa ai requisiti quantitativi e qualitativi, ai termini ed alle condizioni pattuite, come da certificato di regolare esecuzione (prot. n. "),
                         ftext(Prot..prestazione.resa),
                         ftext(") emesso dal Responsabile Unico del Progetto "),
                         ftext(dott.rup), ftext(" "), ftext(RUP), ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATA", fpt.b), ftext(" la dichiarazione resa dall'operatore economico ai sensi della L. 136/2010 in merito alla tracciabilità dei flussi finanziari (c/c dedicato alle commesse pubbliche);")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VERIFICATO", fpt.b), ftext(" che l'IBAN n. "),
                         ftext(Fornitore..IBAN), ftext(" associato all’Operatore Economico in argomento censito nel sistema informativo/contabile dell’Ente, corrisponde a quanto dichiarato dall’Operatore Economico con la dichiarazione resa di cui al punto precedente;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il Documento Unico di Regolarità Contributiva (DURC) che accerta la regolarità della posizione dell’Operatore Economico nei confronti degli Enti individuati dalla normativa vigente alla data odierna;")), style = "Normal")
    if(Importo.senza.IVA.num>=5000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la verifica della posizione dell’Operatore Economico presso l’Agenzia Riscossione Entrate mediante liberatoria di non inadempienza per l’importo della fattura, ai sensi dell’art. 48-bis del DPR n. 602/73, effettuata in data "),
        ftext(format(Sys.Date(), "%d %B %Y")),
        ftext(";")), style = "Normal")
    }
    doc <- doc |>
      body_add_fpar(fpar(ftext("ACCERTATO", fpt.b), ftext(" il diritto del creditore in relazione alla documentazione acquisita;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("DISPONE")), style = "heading 2") |>
      body_add_fpar(fpar(ftext("[scegliere una delle seguenti opzioni e cancellare l'altra]", fpt.i)), style = "Normal") |>
      body_add_fpar(fpar(ftext("la liquidazione e il pagamento della fattura in oggetto per l’importo di "),
                         ftext(Importo.con.IVA),
                         ftext(" pari o inferiore all’importo dell’anticipata citata in premessa, a mezzo bonifico bancario sull’IBAN "),
                         ftext(Fornitore..IBAN),
                         ftext(".")), style = "Normal") |>
      body_add_fpar(fpar(ftext("la liquidazione e il pagamento della fattura in oggetto per l’importo di "),
                         ftext(Importo.con.IVA),
                         ftext(" associata/e alla/e scrittura/e citata/e in premessa per l’importo di "),
                         ftext(Importo.con.IVA),
                         ftext(" e mediante utilizzo delle risorse iscritte a bilancio per il maggior costo per € _____, a mezzo bonifico bancario sull’IBAN "),
                         ftext(Fornitore..IBAN),
                         ftext(".")), style = "Normal") |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext("Controllo di regolarità amministrativa e contabile")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("Responsabile della Gestione e della Compliance amministrativo contabile (RGC)")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(Sig.ra Concetta Mottura)")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext(firma.RSS)), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    
    # if(Importo.senza.IVA.num>=5000){
    #   doc <- doc |>
    #     body_add_par("") |>
    #     body_add_par("") |>
    #     body_add_par("") |>
    #     body_add_fpar(fpar(ftext("VISTO RGC", fpt.b), ftext(" in seguito alla verifica della posizione dell’Operatore Economico presso l’Agenzia Riscossione Entrate mediante liberatoria di non inadempienza per l’importo della fattura, ai sensi dell’art. 48-bis del DPR n. 602/73;")), style = "Firma 1") |>
    #     body_add_fpar(fpar(ftext("(Sig.ra Concetta Mottura)")), style = "Firma 1")
    # }

    print(doc, target = paste0(pre.nome.file, "11 Provv. liquidazione.docx"))
    cat("

    Documento generato: '11 Provv. liquidazione'")

    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Fornitore, Fornitore..Codice.terzo.SIGLA, CIG, CUP, Anticipata, Voce.di.spesa, Fornitore..IBAN, Prot..prestazione.resa)
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

    if(fornitore.uscente=="vero"){
      cat(paste0(frase1, frase2, frase3.1, frase3.2, frase3.3, frase3.4, frase4))
      if(blocco.rota=="vero"){
        stop("Non è possibile continuare. Apportare le modifiche in FluOr come indicato sopra e, poi, generare nuovamente i documenti dopo aver scaricato Ordini.csv.\n")
      }else{
        cat("E' possibile continuare. Premere INVIO per proseguire\n")
        readline()
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
    # if(PNRR!="No"){
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
        headers_replace_text_at_bkm(bookmark = "bookmark_headers", toupper(Progetto.int))
      # file.remove("tmp.docx")
      # file.remove(logo)
    # }else{
    #   doc <- doc |>
    #     headers_replace_text_at_bkm("bookmark_headers", sede1)
    #   if(sede=="TOsi"){
    #     doc <- doc |>
    #       headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    #   }
    #   file.remove("tmp.docx")
    # }
    
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
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2")
    
        if(Richiedente!=Responsabile.progetto){
      doc <- doc |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext(Dott.resp), ftext(" "), ftext(Responsabile.progetto)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(responsabile del progetto e titolare dei fondi)")), style = "Firma 2")
    }
    
    doc <- doc |>
      cursor_bookmark("bookmark_relazione") |>
      body_remove() |>
      cursor_backward()
    
    if(CPV..CPV=="22120000-7"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Una ricerca condotta nell’ambito delle attività del progetto "),
                           ftext(Progetto),
                           ftext(" è stata completata e convogliata in un articolo scientifico scritto da __________ e intitolato “__________”.")), style = "Relazione") |>
        body_add_fpar(fpar(ftext("Indagine di mercato")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("Un’indagine delle riviste scientifiche più adatte a questo articolo per i temi trattati e che abbiamo elevato impact factor, rientrino nel primo quartile (Q1) nel settore "),
                           ftext("Plant Sciences", fpt.i),
                           ftext(" e abbiano prezzi competitivi per pubblicazioni Open Access CC-BY ha portato all’individuazione della rivista __________ pubblicata da "),
                           ftext(Fornitore),
                           ftext(".")), style = "Relazione") |>
        body_add_fpar(fpar(ftext("Dopo peer review, l’articolo è stato ora accettato per la pubblicazione.")), style = "Relazione") |>
        body_add_fpar(fpar(ftext("L’operatore economico è, quindi, "),
                           ftext(Fornitore),
                           ftext(", che offre il servizio di pubblicazione Open Access al costo di "),
                           ftext(Importo.senza.IVA),
                           ftext(" IVA esclusa. Tale fornitore risulta, inoltre, in possesso di esperienze pregresse con altre pubbliche amministrazioni italiane ed è iscritto al MePA.")), style = "Relazione")
      
      if(Rotazione.fornitore!="Non è il contraente uscente"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("L’operatore economico individuato risulta essere contraente uscente. Tuttavia, si chiede l’affidamento all’operatore economico individuato in deroga al principio di rotazione per le seguenti motivazioni, ai sensi dell'art. 49, comma 4 del Codice:")), style = "Relazione") |>
          body_add_fpar(fpar(ftext("struttura del mercato, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("effettiva assenza di alternative, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("accurata esecuzione del precedente contratto, quale __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("trattasi di beni specifici prodotti esclusivamente dal fornitore individuato e funzionali all’attività di ricerca, che richiede continuità e ripetibilità di protocolli operativi specifici;")), style = "Elenco punto")
        if(Importo.senza.IVA.num<5000){
          doc <- doc |>
            body_add_fpar(fpar(ftext("l’importo dell’affidamento è inferiore a euro 5.000,00 (ai sensi dell’art. 49, comma 6, del Codice).")), style = "Elenco punto")
        }
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("Conclusioni")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("In seguito all’accettazione dell’articolo per la pubblicazione, sul sito della rivista viene mostrato il costo, che è pari a "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA. Si richiede, pertanto, l’attivazione dell’idoneo procedimento finalizzato all’acquisizione del servizio in oggetto.")), style = "Relazione")
    }else{
      if(Tipo.acquisizione=="Beni"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Per le attività di ricerca previste nel progetto "),
                             ftext(Progetto),
                             ftext(" è necessaria l'acquisizione "),
                             ftext(della.fornitura),
                             ftext(" di “"),
                             ftext(Prodotto),
                             ftext("”, come dettagliato nella richiesta d'acquisto.")), style = "Relazione")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("Per le attività di ricerca previste nel progetto "),
                             ftext(Progetto),
                             ftext(" è necessaria l'acquisizione "),
                             ftext(della.fornitura),
                             ftext(" di “"),
                             ftext(Prodotto),
                             ftext("” con le seguenti caratteristiche:")), style = "Relazione") |>
          body_add_fpar(fpar(ftext("__________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("__________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("__________;")), style = "Elenco punto")
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("Indagine di mercato")), style = "heading 2")
      
      if(Scelta.fornitore=="Singolo preventivo"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("In seguito ad un’accurata valutazione del mercato è stato acquisito un singolo preventivo, allegato alla presente, per la seguente motivazione: ___________.")), style = "Relazione")
      }else{
        doc <- doc |>
          body_add_fpar(fpar(ftext("In seguito ad un’accurata indagine informale di mercato, con la quale sono stati acquisiti n° ___ preventivi, allegati alla presente, è stato individuato l’operatore economico "),
                             ftext(Fornitore),
                             ftext(" quale potenziale affidatario "),
                             ftext(della.fornitura),
                             ftext(" per le seguenti motivazioni: ___________.")), style = "Relazione")
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("L’operatore economico "),
                           ftext(Fornitore),
                           ftext(" ci ha inviato un preventivo rispondente esattamente alle nostre richieste ed esigenze sia dal punto di vista delle caratteristiche tecniche che dei tempi di consegna, che dal punto di vista del prezzo rispondente agli standard di mercato e con tutte le garanzie richieste sui prodotti. Tale fornitore risulta inoltre in possesso delle esperienze pregresse idonee all’esecuzione della prestazione contrattuale, quali altre forniture simili a pubbliche amministrazioni compreso il CNR.")), style = "Relazione")
      if(Rotazione.fornitore!="Non è il contraente uscente"){
        doc <- doc |>
          body_add_fpar(fpar(ftext("L’operatore economico individuato risulta essere contraente uscente. Tuttavia, si chiede l’affidamento all’operatore economico individuato in deroga al principio di rotazione per le seguenti motivazioni, ai sensi dell'art. 49, comma 4 del Codice:")), style = "Relazione") |>
          body_add_fpar(fpar(ftext("struttura del mercato, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("effettiva assenza di alternative, poichè __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("accurata esecuzione del precedente contratto, quale __________;")), style = "Elenco punto") |>
          body_add_fpar(fpar(ftext("trattasi di beni specifici prodotti esclusivamente dal fornitore individuato e funzionali all’attività di ricerca, che richiede continuità e ripetibilità di protocolli operativi specifici;")), style = "Elenco punto")
        if(Importo.senza.IVA.num<5000){
          doc <- doc |>
            body_add_fpar(fpar(ftext("l’importo dell’affidamento è inferiore a euro 5.000,00 (ai sensi dell’art. 49, comma 6, del Codice).")), style = "Elenco punto")
        }
      }
      doc <- doc |>
        body_add_fpar(fpar(ftext("Conclusioni")), style = "heading 2") |>
        body_add_fpar(fpar(ftext("Da contatti informali, cui è seguita una quotazione budgetaria, il costo massimo omnicomprensivo atteso per l’acquisizione è pari a "),
                           ftext(Importo.senza.IVA),
                           ftext(" oltre IVA. Si richiede, pertanto, l’attivazione dell’idoneo procedimento finalizzato all’acquisizione "),
                           ftext(della.fornitura),
                           ftext(" in oggetto.")), style = "Relazione")
    }
    doc <- doc |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(sede1), ftext(", "), ftext(da))) |>
      body_add_par("") |>
      body_add_fpar(fpar(ftext(Dott.ric), ftext(" "), ftext(Richiedente)), style = "Firma 2") |>
      body_add_fpar(fpar(ftext(firma.RAS)), style = "Firma 2")
    
    if(Richiedente!=Responsabile.progetto){
      doc <- doc |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_par("") |>
        body_add_fpar(fpar(ftext(Dott.resp), ftext(" "), ftext(Responsabile.progetto)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(responsabile del progetto e titolare dei fondi)")), style = "Firma 2")
    }
    
    print(doc, target = paste0(pre.nome.file, "1 RAS.docx"))
    
    cat("

    Documento generato: '1 RAS'")
    
    ## Dich. Ass. RICH ----
    # if(PNRR!="No"){
      download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      # download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      # file.remove("tmp.docx")
      # file.remove(logo)
    # }else{
    #   doc <- doc.ras |>
    #     headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
    #   if(sede=="TOsi"){
    #     doc <- doc |>
    #       headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
    #   }
    #   file.remove("tmp.docx")
    # }
    
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
                         ftext(" all'operatore economico "),
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
      body_add_fpar(fpar(ftext(sede1), ftext(", "),ftext(da)), style = "Normal") |>
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
    cat("

    Documento generato: '4.1 Dichiarazione assenza conflitto RICH'")
    
    ## Dich. Ass. RESP ----
    if(Richiedente!=Responsabile.progetto){
      # if(PNRR!="No"){
        # download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- read_docx("tmp.docx")
        # download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
        doc <- doc |>
          footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
        # file.remove("tmp.docx")
        # file.remove(logo)
      # }else{
      #   doc <- doc.ras |>
      #     headers_replace_all_text("CAMPO.Sede.Secondaria", sede1, only_at_cursor = TRUE)
      #   if(sede=="TOsi"){
      #     doc <- doc |>
      #       headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
      #   }
      #   file.remove("tmp.docx")
      # }
      
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
        body_add_fpar(fpar(ftext(sede1), ftext(", "),ftext(da)), style = "Normal") |>
        body_add_fpar(fpar(ftext("")), style = "Normal") |>
        body_add_fpar(fpar("Il titolare dei fondi e responsabile del progetto", run_footnote(x=block_list(fpar(ftext(" Il dichiarante deve firmare con firma digitale qualificata oppure allegando copia fotostatica del documento di identità, in corso di validità (art. 38 del D.P.R. n° 445/2000 e s.m.i.).", fp_text_lite(italic = TRUE, font.size = 7)))), prop=fp_text_lite(vertical.align = "superscript"))), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("("),
                           ftext(dott.resp),
                           ftext(" "),
                           ftext(Responsabile.progetto),
                           ftext(")")), style = "Firma 2") |>
        body_add_break()
      print(doc, target = paste0(pre.nome.file, "4.2 Dichiarazione assenza conflitto RESP.docx"))
      cat("

    Documento generato: '4.2 Dichiarazione assenza conflitto RESP'")
    }
    file.remove("tmp.docx")
    file.remove(logo)
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore)
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
    if(fornitore.uscente=="vero"){
      cat(paste0(frase1, frase2, frase3.1, frase3.2, frase3.3, frase3.4, frase4))
      if(blocco.rota=="vero"){
        stop("Non è possibile continuare. Apportare le modifiche in FluOr come indicato sopra e, poi, generare nuovamente i documenti dopo aver scaricato Ordini.csv.\n")
      }else{
        cat("E' possibile continuare. Premere INVIO per proseguire\n")
        readline()
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
                         ftext(": NOMINA DEL RESPONSABILE UNICO DEL PROGETTO AI SENSI DELL’ART. 15 E DELL’ALLEGATO I.2 DEL DECRETO LEGISLATIVO 31 MARZO 2023 N. 36 PER L’AFFIDAMENTO DIRETTO "),
                         ftext(toupper(della.fornitura)),
                         ftext(" DI “"),
                         ftext(toupper(Prodotto), fpt.b),
                         ftext("”"),
                         ftext(", ORDINE "),
                         ftext(sede, fpt.b),
                         ftext(" "),
                         ftext(ordine, fpt.b),
                         ftext(y, fpt.b),
                         ftext(", NELL'AMBITO DEL "),
                         ftext(toupper(Progetto.int)),
                         ftext(".")), style = "Normal") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      cursor_reach("CAMPO.NOMINE") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")
    
    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento n. 31 del Direttore Generale del CNR prot. n. 54747 del 25/2/2025 di attribuzione dell'incarico di Responsabile della Gestione e Compliance amministrativo-contabile (RGC) dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) alla sig.ra Concetta Mottura per il periodo dall’1/3/2025 al 29/2/2028;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
      #                   ftext(nomina.RAMM)), style = "Normal") |>
      #body_add_fpar(fpar(ftext("VISTA", fpt.b), ftext(" la delega del Segretario Amministrativo dell’IPSP al Responsabile Amministrativo della "),
      #                   ftext(sede2), ftext(" dell’IPSP prot. 153859 dell’8/5/2024 per il periodo dall’8/5/2024 al 31/12/2024 ad effettuare il controllo interno di regolarità amministrativa e copertura finanziaria per gli affidamenti diretti ed apporre il visto sulla “Decisione di contrattare” prevista dall’art. 32 del Regolamento di Amministrazione Contabilità e Finanza (RACF) del Consiglio Nazionale delle Ricerche, emanato con provvedimento della Presidente CNR n. 201 del 23 dicembre 2024, in vigore dal 1° gennaio 2025;")), style = "Normal") |>
      cursor_reach("CAMPO.DECRETO") |>
      body_remove() |>
      cursor_backward() |>
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
      cursor_reach("CAMPO.DISPONIBILITA") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b), ftext(" la disponibilità economico-finanziaria per la copertura degli oneri derivanti dall'acquisizione "),
                         ftext(della.fornitura),ftext(" di “"),
                         ftext(Prodotto),
                         ftext("” sui fondi del progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce di costo CO.AN "),
                         ftext(Voce.di.spesa), ftext(";")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATA", fpt.b), ftext(", pertanto, la necessità di procedere:")), style = "Normal")

    if(Supporto.RUP!=trattini){
      doc <- doc |>
        body_add_fpar(fpar(ftext("alla nomina del responsabile unico del progetto (RUP) per la programmazione, progettazione, affidamento e l’esecuzione dell’affidamento "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto), ftext("”;")), style = "Elenco liv2")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("alla nomina del responsabile unico del progetto (RUP) per la programmazione, progettazione, affidamento e l’esecuzione dell’affidamento "),
                           ftext(della.fornitura), ftext(" di “"),
                           ftext(Prodotto), ftext("”, nonché alla nomina del supporto al RUP;")), style = "Elenco liv2")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("alla scrittura anticipata CO.AN inerente le somme necessarie indicate nella richiesta d’acquisto prot. n. "),
                         ftext(Prot..RAS), ftext(";")), style = "Elenco liv2") |>
      cursor_bookmark("bookmark_nomina_rup") |>
      body_remove() |>
      cursor_backward()
      
    if(RUP!=RSS.nome){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI NOMINARE", fpt.b), ftext(" "), ftext(il.dott.rup), ftext(" "), ftext(RUP, fpt.b),
                           ftext(" Responsabile Unico del Progetto (RUP) che, ai sensi dell'art. 15 del Codice, dovrà:")), style = "Elenco liv1")
    }else{
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI ASSUMERE", fpt.b), ftext(" il ruolo di Responsabile Unico del Progetto (RUP) che, ai sensi dell'art. 15 del Codice, dovrà:")), style = "Elenco liv1")
    }

    doc <- doc |>
      cursor_bookmark("bookmark_supporto_rup")
    
      if(Supporto.RUP!=trattini){
        doc <- doc |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("DI INDIVIDUARE", fpt.b), ftext(" ai sensi dell'art. 15, comma 6, del Codice, "),
                             ftext(il.dott.sup), ftext(" "), ftext(Supporto.RUP, fpt.b),
                             ftext(" in qualità di supporto al RUP, fermo restando i compiti e le mansioni a cui gli stessi sono già assegnati;")), style = "Elenco liv1")
      }else{
        doc <- doc |>
          body_remove()
      }
      
    doc <- doc |>
      body_replace_text_at_bkm(bookmark = "bookmark_A1_fornitura", della.fornitura) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A1", formatC(Importo.senza.IVA.num - Manodopera.num - Oneri.sicurezza.num, digits=2, format="f", decimal.mark=",", big.mark = ".")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A2", formatC(Manodopera.num, digits=2, format="f", decimal.mark=",", big.mark = ".")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A3", formatC(Oneri.sicurezza.num, digits=2, format="f", decimal.mark=",", big.mark = ".")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_A", formatC(Importo.senza.IVA.num, digits=2, format="f", decimal.mark=",", big.mark = ".")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_AB", formatC(Importo.senza.IVA.num, digits=2, format="f", decimal.mark=",", big.mark = ".")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_C4", formatC(IVA.num, digits=2, format="f", decimal.mark=",", big.mark = "."))
    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
      body_replace_text_at_bkm(bookmark = "bookmark_C", formatC(IVA.num, digits=2, format="f", decimal.mark=",", big.mark = ".")) |>
      body_replace_text_at_bkm(bookmark = "bookmark_ABC", formatC(Importo.senza.IVA.num + IVA.num, digits=2, format="f", decimal.mark=",", big.mark = "."))
    }else{
      doc <- doc |>
        body_replace_text_at_bkm(bookmark = "bookmark_C2", "35") |>
        body_replace_text_at_bkm(bookmark = "bookmark_C", formatC(35 + IVA.num, digits=2, format="f", decimal.mark=",")) |>
        body_replace_text_at_bkm(bookmark = "bookmark_ABC", formatC(35 + Importo.senza.IVA.num + IVA.num, digits=2, format="f", decimal.mark=","))
    }
    
    doc <- doc |>
      cursor_bookmark("bookmark_confermare") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("DI CONFERMARE", fpt.b), ftext(" la registrazione sul sistema contabile della seguente scrittura anticipata n. "),
                         ftext(Anticipata),
                         ftext(" di "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre iva sul progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce di costo CO.AN "),
                         ftext(Voce.di.spesa),                   
                         ftext(";")), style = "Elenco liv1")
    
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI CONFERMARE", fpt.b), ftext(" la registrazione sul sistema contabile della seguente scrittura anticipata _____ del _____ di 35,00 € sul progetto"),
                           ftext(Progetto.cup),
                           ftext(", voce di costo CO.AN 13096 per la contribuzione ANAC;")), style = "Elenco liv1")
    }
    
    if(CCNL!="Non applicabile"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI DARE ATTO ", fpt.b), ftext("che:")), style = "Elenco liv1") |>
        body_add_fpar(fpar(ftext("ai sensi dell’art.11 del Codice l’O.E. affidatario sarà tenuto ad applicare il seguente CCNL territoriale individuato dalla S.A.: "),
        ftext(CCNL),
        ftext(";")), style = "Elenco liv2") |>
        body_add_fpar(fpar(ftext("i costi della manodopera indicati nel quadro economico sopra riportato sono stati calcolati sulla base delle tariffe orarie previste per il CCNL individuato al punto che precede;")), style = "Elenco liv2")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("DI RENDERE", fpt.b), ftext(" consultabile il presente atto sulla piattaforma telematica di negoziazione da parte dell’O.E. invitato a presentare offerta, unitamente:")), style = "Elenco liv1") |>
      body_add_fpar(fpar(ftext("alla richiesta d’acquisto prot. n. "),
                         ftext(Prot..RAS), ftext(";")), style = "Elenco liv2") |>
      body_add_fpar(fpar(ftext("alle condizioni generali d'acquisto da sottoscrivere successivamente;")), style = "Elenco liv2") |>
      cursor_reach("CAMPO.FIRMA") |>
      body_remove() |>
      # body_add_fpar(fpar(ftext("Controllo di regolarità contabile")), style = "Firma 1") |>
      # body_add_fpar(fpar(ftext("Responsabile della Gestione e della Compliance amministrativo contabile (RGC)")), style = "Firma 1") |>
      # body_add_fpar(fpar(ftext("(Sig.ra Concetta Mottura)")), style = "Firma 1") |>
      #body_add_par("Visto di regolarità contabile", style = "Firma 1") |>
      #body_add_par(resp.segr, style = "Firma 1") |>
      #body_add_fpar(fpar(ftext("("), ftext(RAMM), ftext(")")), style = "Firma 1") |>
      #body_add_par("La segretaria amministrativa", style = "Firma 1") |>
      #body_add_fpar(fpar(ftext("(sig.ra Concetta Mottura)")), style = "Firma 1") |>
      #body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    print(doc, target = paste0(pre.nome.file, "2 Nomina RUP.docx"))
    
    cat("

    Documento generato: '2 Nomina RUP'")
    
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
                         ftext(", all'operatore economico "),
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

    Documento generato: '4.3 Dichiarazione assenza conflitto RSS'")
    
    ## Dich. Ass. RUP ----
    #download.file(paste(lnk, "Dich_conf.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    #download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
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
                         ftext(", all'operatore economico "),
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
    cat("

    Documento generato: '4.4 Dichiarazione assenza conflitto RUP'")
    
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
                           ftext(", all'operatore economico "),
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
        body_add_fpar(fpar(ftext(sede1), ftext(", "),ftext(da)), style = "Normal") |>
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

    Documento generato: '4.5 Dichiarazione assenza conflitto SUP'")
    }
    
    ## Patto integrità ----
    if(Fornitore..Rappresentante.legale!=trattini){
      download.file(paste(lnk, "Patto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      #download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      
      if(sede!="TOsi"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_rss", paste0(", che delega alla firma ", paste0(tolower(substr(firma.RSS, 1, 1)),substr(firma.RSS, 2, nchar(firma.RSS))),  " ", RSS))
      }
      
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
      
      # file.remove("tmp.docx")
      # file.remove(logo)
      print(doc, target = paste0(pre.nome.file, "3.3 Patto di integrità.docx"))
      cat("

    Documento generato: '3.3 Patto di integrità'")
    }
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore, RUP, Prot..RAS)
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
                  #"CCNL",
                  "bollo")
      docuOE_ext <- c("3.1 Dichiarazione possesso requisiti di partecipazione e di qualificazione infra 40k",
                      "3.1 Dichiarazione possesso requisiti di qualificazione oltre 40k",
                      "3.4 Dichiarazione DPCM 187 1991",
                      "3.6 Dichiarazione titolare effettivo",
                      "3.8 Dichiarazione ausiliaria",
                      #"3.9 Comprova equivalenza tutele CCNL",
                      "3.11 Comprova imposta di bollo")
      if(Importo.senza.IVA.num<40000){
        docuOE <- docuOE[c(-2,-6)]
        docuOE_ext <- docuOE_ext[c(-2,-6)]
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
    
    if(CCNL!="Non applicabile"){
    ## Manodopera ----
      download.file(paste(lnk, "Manodopera.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      file.remove("logo")
      
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
        body_replace_text_at_bkm("bookmark_intestazione", int.docoe) 
        
      print(doc, target = paste0(pre.nome.file, "3.9 Costi manodopera.docx"))
    }
    
    ## CC dedicato ----
    download.file(paste(lnk, "cc_dedicato.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    print(doc, target = paste0(pre.nome.file, "3.2 Comunicazione conto corrente dedicato.docx"))
    
    ## Patto d'integrità ----
    download.file(paste(lnk, "Patto.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      
      if(sede!="TOsi"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_rss", paste0(", che delega alla firma ", paste0(tolower(substr(firma.RSS, 1, 1)),substr(firma.RSS, 2, nchar(firma.RSS))),  " ", RSS))
      }
      
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
    
    
    ## Declaration on honour ----
    if(Fornitore..Nazione=="Estera"){
      download.file(paste(lnk, "Honour.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      print(doc, target = paste0(pre.nome.file, "3.3 Declaration on honour.docx"))
      cat("

    Documento generato: '3.4 Declaration on honour'")
    }
    
    ## CAM ----
    if(substr(CPV,1,3)=="301" | substr(CPV,1,3)=="302" | CPV==trattini){
      download.file(paste(lnk, "CAM.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
      print(doc, target = paste0(pre.nome.file, "3.10 Documentazione rispetto CAM.docx"))
    }
    
    ## DNSH ----
    if(Inventariabile=='Non inventariabile'){
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
    }
    
    if(Inventariabile=='Inventariabile'){
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
    }
    
    if(Tipo.acquisizione=='Beni' & Inventariabile=='Non inventariabile'){
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
    }
    
    if(Tipo.acquisizione=='Servizi'){
      download.file(paste(lnk, "DNSH_26.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
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
      print(doc, target = paste0(pre.nome.file, "3.5 Scheda DNSH servizi di ricerca.docx"))
    }
    
    ## Condizioni d'acquisto ----
    if(Fornitore..Nazione=="Italiana"){
    download.file(paste(lnk, "Condizioni.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      headers_replace_text_at_bkm(bookmark = "bookmark_headers", toupper(Progetto.int))
    
    if(Inventariabile=="Inventariabile"){
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_durata", "la fornitura dovrà essere consegnata e installata entro 6 mesi")
    }
    if(Tipo.acquisizione=="Servizi"){
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_durata", "il servizio dovrà essere svolto entro 6 mesi")
    }
    
    doc <- doc |>
      body_replace_text_at_bkm("bookmark_dicitura_fattura", dicitura.fattura)
    
    if(Tipo.acquisizione=='Beni'){
      doc <- doc |>
        cursor_bookmark("bookmark_conformita") |>
        body_remove() |>
        cursor_backward() |>
        #body_replace_all_text("CAMPO.CONFORMITA", "", only_at_cursor = TRUE) |>
        body_add_fpar(fpar(ftext("Verifica di conformità", fpt.b), ftext(": la presente fornitura è soggetta a verifica di conformità da effettuarsi, secondo quanto previsto dall’art. 116 e nell’Allegato II.14 del codice dei contratti entro 2 mesi. A seguito della verifica di conformità si procede al pagamento della rata di saldo e, se prevista, allo svincolo della cauzione.")), style = "Elenco punto")
    }else{
      doc <- doc |>
        cursor_bookmark("bookmark_conformita") |>
        body_remove() |>
        cursor_backward() |>
        #body_replace_all_text("CAMPO.CONFORMITA", "", only_at_cursor = TRUE) |>
        body_add_fpar(fpar(ftext("Verifica di regolare esecuzione", fpt.b), ftext(": la stazione appaltante, per il tramite del RUP, emette il certificato di regolare esecuzione, secondo le modalità indicate nell'Allegato II.14 al codice dei contratti pubblici, entro 6 mesi. A seguito dell’emissione del certificato di regolare esecuzione si procede al pagamento della rata di saldo e, se prevista, allo svincolo della cauzione.")), style = "Elenco punto")
    } 
    
    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Clausola risolutiva espressa", fpt.b), ftext(": il CNR ha diritto di risolvere il contratto/ordine in caso di accertamento della carenza dei requisiti di partecipazione. Per la risoluzione del contratto trovano applicazione l’art. 122 del d.lgs. 36/2023, nonché gli articoli 1453 e ss. del Codice Civile. Il CNR darà formale comunicazione della risoluzione al fornitore, con divieto di procedere al pagamento dei corrispettivi, se non nei limiti delle prestazioni già eseguite.")), style = "Elenco punto")
    }
    
    print(doc, target = paste0(pre.nome.file, "3.11 Condizioni generali di acquisto.docx"))
    }else{
      download.file(paste(lnk, "Condizioni_eng.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      doc <- doc |>
        footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
        headers_replace_text_at_bkm(bookmark = "bookmark_headers", toupper(Progetto.int))
      
      if(Inventariabile=="Inventariabile"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_durata_eng", "the supply must be delivered and installed within 6 months")
      }
      if(Tipo.acquisizione=="Servizi"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_durata_eng", "the service must be carried out within 6 months")
      }
      
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_dicitura_fattura_eng", dicitura.fattura)
      
      if(Tipo.acquisizione=='Beni'){
        doc <- doc |>
          cursor_bookmark("bookmark_conformita_eng") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Compliance check", fpt.b), ftext(": this supply is subject to a conformity check to be carried out, as per art. 116 and Annex II.14 of the contracts code within 2 months. Following the conformity check, the balance instalment will be paid and, if applicable, the deposit will be released.")), style = "Elenco punto 2")      }else{
        doc <- doc |>
          cursor_bookmark("bookmark_conformita_eng") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Verification of proper execution", fpt.b), ftext(": the contracting authority, through the RUP, issues the certificate of proper execution, according to the methods indicated in Annex II.14 to the public contracts code, within 6 months. Following the issuance of the certificate of proper execution, the payment of the balance instalment and, if applicable, the release of the security deposit shall take place.")), style = "Elenco punto 2")
      } 
      
      if(Importo.senza.IVA.num<40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Express termination clause", fpt.b), ftext(": CNR has the right to terminate the contract/order in the event of a lack of participation requirements being ascertained. For the termination of the contract, art. 122 of Legislative Decree 36/2023, as well as articles 1453 et seq. of the Civil Code, apply. CNR will formally communicate the termination to the supplier, with a ban on proceeding with the payment of the fees, except within the limits of the services already performed.")), style = "Elenco punto 2")
      }
    }
    
    print(doc, target = paste0(pre.nome.file, "3.11 Condizioni generali di acquisto.docx"))
    
    ## Privacy ----
    if(Fornitore..Nazione=="Italiana"){
    download.file(paste(lnk, "Privacy.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    file.remove("tmp.docx")
    
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      cursor_bookmark("bookmark_oggetto") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("La presente informativa descrive le misure di tutela riguardo al trattamento dei dati personali destinata ai fornitori di beni e/o servizi, nell’ambito dell’affidamento diretto "),
                         ftext(della.fornitura),
                         ftext(" di “"),
                         ftext(Prodotto, fpt.b),
                         ftext("”, ai sensi dell’articolo 13 del Regolamento UE 2016/679 in materia di protezione dei dati personali (di seguito, per brevità, GDPR).")), style = "Normal")
      
    print(doc, target = paste0(pre.nome.file, "3.12 Informativa privacy.docx"))
    }else{
      download.file(paste(lnk, "Privacy_eng.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
      doc <- read_docx("tmp.docx")
      file.remove("tmp.docx")
      
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_oggetto_eng", Prodotto)
      print(doc, target = paste0(pre.nome.file, "3.12 Privacy policy.docx"))
    }
    
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

    Documenti generati: 'Autocertificazioni dell'operatore economico'
    Documento generato: '3.11 Condizioni generali di acquisto'
    Documento generato: '3.12 Informativa privacy'
        ")
  }
  
  # AI PNRR ----
  ai.pnrr <- function(){

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
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento prot. n. "),
                         ftext(Prot..nomina.RUP),
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
                         ftext("), un’offerta ritenuta congrua, corredata dalle dichiarazioni sostitutive richieste, in merito al possesso dei requisiti prescritti dalla stazione appaltante, corrispondente al preventivo precedentemente acquisito e agli atti;")), style = "Normal")
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
      body_replace_text_at_bkm(bookmark = "bookmark_oe", paste0(Fornitore, " (P.IVA ", Fornitore..P.IVA, ", soggetto U-Gov ", Fornitore..Codice.terzo.SIGLA, ")"))
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
    
    cat("

    Documento generato: '5 Atto istruttorio'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Fornitore, RUP, Prot..RAS, Pagina.web, Fornitore..Codice.terzo.SIGLA, RDO)
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
    
  }
  
  # DaC PNRR ----
  dac.pnrr <- function(){
    
    if(fornitore.uscente=="vero"){
      cat(paste0(frase1, frase2, frase3.1, frase3.2, frase3.3, frase3.4, frase4))
      if(blocco.rota=="vero"){
        stop("Non è possibile continuare. Apportare le modifiche in FluOr come indicato sopra e, poi, generare nuovamente i documenti dopo aver scaricato Ordini.csv.\n")
      }else{
        cat("E' possibile continuare. Premere INVIO per proseguire\n")
        readline()
      }
    }
    
    download.file(paste(lnk, "DaC.docx", sep=""), destfile = "tmp.docx", method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- read_docx("tmp.docx")
    download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm"))
    file.remove("tmp.docx")
    file.remove(logo)
    
    if(sede=="TOsi"){
      doc <- doc |>
        headers_replace_all_text("Secondaria", "Istituzionale", only_at_cursor = TRUE)
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
                         ftext(" ("),
                         ftext(toupper(Pagina.web)),
                         ftext("), NELL'AMBITO DEL "),
                         ftext(toupper(Progetto.int)),
                         ftext(".")), style = "Normal") |>
      body_add_par(firma.RSS, style = "heading 2") |>
      cursor_reach("CAMPO.NOMINE") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, con cui al dott. Francesco Di Serio è stato attribuito l’incarico di Direttore dell’IPSP del Consiglio Nazionale delle Ricerche a decorrere dal giorno 1/5/2024 per quattro anni;")), style = "Normal")
    
    if(sede!="TOsi"){
      doc <- doc |>
        body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. "),
                           ftext(nomina.RSS), ftext(", il quale è autorizzato ad intraprendere ogni atto necessario per procedere agli acquisti di beni e servizi, nonché esecuzione di lavori, fino all’importo complessivo € 15.000,00 (IVA esclusa);")), style = "Normal")
    }
    
    doc <- doc |>
      body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento n. 31 del Direttore Generale del CNR prot. n. 54747 del 25/2/2025 di attribuzione dell'incarico di Responsabile della Gestione e Compliance amministrativo-contabile (RGC) dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) alla sig.ra Concetta Mottura per il periodo dall’1/3/2025 al 29/2/2028;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore dell’IPSP prot. 146189 del 2/5/2024 di nomina della sig.ra Concetta Mottura quale Segretario Amministrativo dell’IPSP (con sede istituzionale a Torino, centro di spesa 121) per il periodo dall’1/5/2024 fino al 31/12/2024;")), style = "Normal") |>
      # body_add_fpar(fpar(ftext("VISTO", fpt.b), ftext(" il provvedimento del Direttore Generale (prot. 502457 del 18/12/2024) di proroga operativa delle funzioni di Segretario Amministrativo abilitato alla firma degli ordinativi finanziari e del controllo interno di regolarità amministrativo-contabile delle strutture dell’Ente nelle more del conferimento delle nomine a Responsabili della Gestione e della Compliance amministrativo contabile (RGC);")), style = "Normal") |>
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
                         ftext("”, nell’ambito delle attività previste dal progetto "),
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
                         ftext(tolower(Prot..nomina.RUP)),
                         ftext(", con il quale è "),
                         ftext(nominato),
                         ftext(" "),
                         ftext(il.dott.rup),
                         ftext(" "),
                         ftext(RUP),
                         ftext(" quale Responsabile Unico del Progetto (RUP) ai sensi dell’art. 15 del Codice;")), style = "Normal") |>
      body_add_fpar(fpar(ftext("CONSIDERATO", fpt.b), ftext(" che l’operatore economico individuato "),
                         ftext(paste0(Fornitore, " (P.IVA ", Fornitore..P.IVA, ")")),
                         ftext(" ha presentato, attraverso la piattaforma telematica di negoziazione, un’offerta ritenuta congrua, corredata dalle dichiarazioni sostitutive richieste, in merito al possesso dei requisiti prescritti dalla S.A., d'importo corrispondente al preventivo precedentemente acquisito e agli atti;")), style = "Normal") |>
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
      body_add_fpar(fpar(ftext("ACCERTATA", fpt.b), ftext(" la disponibilità finanziaria per la copertura della spesa sui fondi del progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce di costo CO.AN "), ftext(Voce.di.spesa), ftext(";")), style = "Normal") |>
    body_add_fpar(fpar(ftext("VALUTATA", fpt.b), ftext(" l’opportunità, in ottemperanza alla suddetta normativa, di procedere ad affidamento diretto all’operatore economico "),
                       ftext(Fornitore),
                       ftext(" (P.IVA "),
                       ftext(Fornitore..P.IVA),
                       ftext("), mediante provvedimento contenente gli elementi essenziali descritti nell'art. 17, comma 2, del Codice, tenuto conto che il medesimo è in possesso di documentate esperienze pregresse idonee all'esecuzione della prestazione contrattuale;")), style = "Normal")
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
                           ftext(", soggetto U-Gov "),
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
                           ftext(", soggetto U-Gov "),
                           ftext(Fornitore..Codice.terzo.SIGLA),
                           ftext("), con sede legale in "),
                           ftext(Fornitore..Sede),
                           ftext(", individuato mediante indagine informale di mercato, che ha presentato il proprio preventivo ammontante a "),
                           ftext(Importo.senza.IVA, fpt.b),
                           ftext(" oltre IVA;")), style = "Elenco numero")
    }
    doc <- doc |>
      cursor_bookmark("bookmark_confermare") |>
      body_remove() |>
      cursor_backward() |>
      body_add_fpar(fpar(ftext("DI CONFERMARE", fpt.b),
                         ftext(" la registrazione sul sistema contabile della seguente scrittura anticipata n. "),
                         ftext(Anticipata),
                         ftext(" di "),
                         ftext(Importo.senza.IVA),
                         ftext(" oltre IVA sul progetto "),
                         ftext(Progetto.cup),
                         ftext(", voce di costo CO.AN "),
                         ftext(Voce.di.spesa),
                         ftext(";")), style = "Elenco numero")
    if(Importo.senza.IVA.num>=40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("DI CONFERMARE", fpt.b),
                           ftext(" la registrazione sul sistema contabile della seguente scrittura anticipata n. _________ di € 35,00, sul progetto "),
                           ftext(Progetto.cup), 
                           ftext(", voce di costo 13096 per la contribuzione ANAC;")), style = "Elenco numero")
    } 
    doc <- doc |>
      cursor_reach("CAMPO.FIRMA") |>
      body_remove() |>
      body_add_fpar(fpar(ftext("Controllo di regolarità contabile")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("Responsabile della Gestione e della Compliance amministrativo contabile (RGC)")), style = "Firma 1") |>
      body_add_fpar(fpar(ftext("(Sig.ra Concetta Mottura)")), style = "Firma 1") |>
      body_add_par("", style = "Normal") |>
      body_add_fpar(fpar(firma.RSS), style = "Firma 2") |>
      body_add_fpar(fpar(ftext("("), ftext(RSS), ftext(")")), style = "Firma 2")
    
    print(doc, target = paste0(pre.nome.file, "6 Decisione a contrattare.docx"))
    
    cat("

    Documento generato: '6 Decisione a contrattare'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, Richiedente, Importo.senza.IVA, Voce.di.spesa, CUP, Responsabile.progetto, Fornitore, RUP, Prot..RAS, Pagina.web, Prot..atto.istruttorio, Anticipata)
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
    download.file(paste(lnk, logo, sep=""), destfile = logo, method = "curl", extra = "--ssl-no-revoke", quiet = TRUE)
    doc <- doc |>
      footers_replace_img_at_bkm(bookmark = "bookmark_footers", external_img(src = logo, width = 3, height = 2, unit = "cm")) |>
      headers_replace_text_at_bkm(bookmark = "bookmark_headers", toupper(Progetto.int))
    file.remove("tmp.docx")
    file.remove(logo)

    doc <- doc |>
      cursor_begin() |>
      body_add_fpar(fpar(ftext("LETTERA D’ORDINE "), ftext(sede), ftext(" "), ftext(ordine), ftext(y)), style = "heading 1", pos = "on") |>
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
      body_replace_all_text("CAMPO.DAC.LDO", paste0(Prot..DaC, " (", Pagina.web, ")"), only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.RDO1") |>
      body_replace_all_text("CAMPO.RDO1", ordine.trattativa.scelta.ldo1, only_at_cursor = TRUE) |>
      cursor_reach("CAMPO.RDO2") |>
      body_replace_all_text("CAMPO.RDO2", as.character(ordine.trattativa.scelta.ldo2), only_at_cursor = TRUE) |>
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
      body_replace_text_at_bkm("bookmark_cuu", CUU) |>
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
    
    if(Inventariabile=="Inventariabile"){
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_durata", "la fornitura dovrà essere consegnata e installata entro 6 mesi")
    }
    if(Tipo.acquisizione=="Servizi"){
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_durata", "il servizio dovrà essere svolto entro 6 mesi")
    }
    
    doc <- doc |>
      body_replace_text_at_bkm("bookmark_dicitura_fattura", dicitura.fattura)
    
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
          body_add_fpar(fpar(ftext("Verifica di regolare esecuzione", fpt.b), ftext(": la stazione appaltante, per il tramite del RUP, emette il certificato di regolare esecuzione, secondo le modalità indicate nell'Allegato II.14 al codice dei contratti pubblici, entro 6 mesi. A seguito dell’emissione del certificato di regolare esecuzione si procede al pagamento della rata di saldo e, se prevista, allo svincolo della cauzione.")), style = "Elenco punto")
      } 
    
    if(Importo.senza.IVA.num<40000){
      doc <- doc |>
        body_add_fpar(fpar(ftext("Clausola risolutiva espressa", fpt.b), ftext(": il CNR ha diritto di risolvere il contratto/ordine in caso di accertamento della carenza dei requisiti di partecipazione. Per la risoluzione del contratto trovano applicazione l’art. 122 del d.lgs. 36/2023, nonché gli articoli 1453 e ss. del Codice Civile. Il CNR darà formale comunicazione della risoluzione al fornitore, con divieto di procedere al pagamento dei corrispettivi, se non nei limiti delle prestazioni già eseguite.")), style = "Elenco punto")
    }
    
    if(Fornitore..Nazione=="Italiana"){
      b <- cursor_reach(doc, "CAMPO.INIZIO.LDO.EN")
      b <- doc$officer_cursor$which
      e <- cursor_end(doc)
      e <- e$officer_cursor$which -6
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
        body_replace_all_text("CAMPO.DAC.LDO", paste0(Prot..DaC.en, " (", Pagina.web, ")"), only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.RDO1") |>
        body_replace_all_text("CAMPO.RDO1", ordine.trattativa.scelta.ldo1, only_at_cursor = TRUE) |>
        cursor_reach("CAMPO.RDO2") |>
        body_replace_all_text("CAMPO.RDO2", as.character(ordine.trattativa.scelta.ldo2), only_at_cursor = TRUE) |>
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
        body_replace_text_at_bkm("bookmark_cuu_en", CUU) |>
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
      
      if(Inventariabile=="Inventariabile"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_durata_eng", "the supply must be delivered and installed within 6 months")
      }
      if(Tipo.acquisizione=="Servizi"){
        doc <- doc |>
          body_replace_text_at_bkm("bookmark_durata_eng", "the service must be carried out within 6 months")
      }
      
      doc <- doc |>
        body_replace_text_at_bkm("bookmark_dicitura_fattura_eng", dicitura.fattura)

      if(Tipo.acquisizione=='Beni'){
        doc <- doc |>
          cursor_bookmark("bookmark_conformita_eng") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Verifica di conformità", fpt.b), ftext(": this supply is subject to a conformity check to be carried out, as per art. 116 and Annex II.14 of the contracts code within 2 months. Following the conformity check, the balance instalment will be paid and, if applicable, the deposit will be released.")), style = "Elenco punto")
      }else{
        doc <- doc |>
          cursor_bookmark("bookmark_conformita_eng") |>
          body_remove() |>
          cursor_backward() |>
          body_add_fpar(fpar(ftext("Verifica di regolare esecuzione", fpt.b), ftext(": the contracting authority, through the RUP, issues the certificate of proper execution, according to the methods indicated in Annex II.14 to the public contracts code, within 6 months. Following the issuance of the certificate of proper execution, the payment of the balance instalment and, if applicable, the release of the security deposit shall take place.")), style = "Elenco punto")
      } 
      
      if(Importo.senza.IVA.num<40000){
        doc <- doc |>
          body_add_fpar(fpar(ftext("Express termination clause", fpt.b), ftext(": CNR has the right to terminate the contract/order in the event of a lack of participation requirements being ascertained. For the termination of the contract, art. 122 of Legislative Decree 36/2023, as well as articles 1453 et seq. of the Civil Code, apply. CNR will formally communicate the termination to the supplier, with a ban on proceeding with the payment of the fees, except within the limits of the services already performed.")), style = "Elenco punto 2")
      }
    }
    print(doc, target = paste0(pre.nome.file, "7 Lettera ordine.docx"))
    
    cat("

    Documento generato: '7 Lettera ordine'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, CIG, RUP, RDO, Fornitore, Importo.senza.IVA, Aliquota.IVA, Pagina.web, Prot..DaC)
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
                         ftext(Fornitore), ftext(" (P.IVA "), ftext(Fornitore..P.IVA), ftext("; soggetto U-Gov "), ftext(Fornitore..Codice.terzo.SIGLA), ftext(");")), style = "Elenco punto") |>
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

    Documento generato: '8 Prestazione resa'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, CIG, RUP, RDO, Fornitore, Importo.senza.IVA, Aliquota.IVA, Pagina.web, Prot..DaC, DDT, Fornitore..Codice.terzo.SIGLA)
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
      body_add_fpar(fpar(ftext("Il sottoscritto dott. Francesco Di Serio, nato a Cava de’ Tirreni (SA) il 29/09/1965, codice fiscale DSRFNC65P29C361R, direttore dell'IPSP a decorrere dal giorno 1/5/2024 per quattro anni in base al provvedimento del Direttore Generale del Consiglio Nazionale delle Ricerche n. 69 prot. AMMCNT-CNR 140496 del 29/4/2024, in relazione all'affidamento diretto "),
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
    
    cat("

    Documento generato: '9 Dichiarazione assenza doppio finanziamento'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Prodotto, Progetto, CUP, Fornitore, Prot..DaC, Pagina.web)
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
        body_add_fpar(fpar(ftext(sottoscritto.resp), ftext(" "), ftext(dott.resp), ftext(" "), ftext(Responsabile.progetto), ftext(", "),
                           ftext(nato.resp), ftext(" "), ftext(Responsabile.progetto..Luogo.di.nascita), ftext(" il "),
                           ftext(Responsabile.progetto..Data.di.nascita), ftext(", codice fiscale "), ftext(Responsabile.progetto..Codice.fiscale), ftext(", in merito allo strumento “"),
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
        body_add_fpar(fpar(ftext(Dott.resp), ftext(" "), ftext(Responsabile.progetto)), style = "Firma 2") |>
        body_add_fpar(fpar(ftext("(responsabile del progetto e titolare dei fondi)")), style = "Firma 2")
      
      print(doc, target = paste0(pre.nome.file, "10 Dichiarazione funzionalità bene.docx"))
      
      cat("

    Documento generato: '10 Dichiarazione funzionalità bene'")
      
      ## Dati mancanti ---
      manca <- dplyr::select(sc, Prodotto, Progetto, CUP, Fornitore, Prot..DaC, Pagina.web)
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
      body_replace_all_text("CAMPO.DOCOE", Prot..DocOE, only_at_cursor = FALSE) |>
      body_replace_all_text("CAMPO.CONF", Prot..conflitto.interesse, only_at_cursor = FALSE) |>
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

    Documento generato: '11 Checklist'")
    
    ## Dati mancanti ---
    manca <- dplyr::select(sc, Importo.senza.IVA, Prodotto, Fornitore, Prot..DaC, Prot..atto.istruttorio, Prot..conflitto.interesse, Prot..lettera.ordine, Prot..doppio.finanziamento, Pagina.web, CIG, CUP, RDO)
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
         
        | Ordine N° ", ordine, " '", Prodotto, "'", sep="")
    cat("
        | Fornitore: ", Fornitore, sep="")
    cat("
        | Progetto: ", Progetto, sep="")
    cat("
         ___________________________")
    if(PNRR=="No"){
    cat("

    Che documento vuoi generare?
    (per singoli documenti digitare la cifra in parentesi, ad es. 3.1 per Autocertificazioni operatore economico)
      1: RAS, Richiesta pagina web
      2: Nomina RUP, Provvedimento registrazione anticipata
      3: Autocertificazioni operatore economico (.1), Atto istruttorio e Comunicazione CIG (.2),
         Decisione a contrattare (.3), Lettera d'ordine (.4), Certificato di regolare esecuzione (.5),
         Provvedimento di liquidazione (.6)

")
      
      inpt <- readline()
      if(inpt==1){cat("\014");ras();pag()}
      if(inpt==2){cat("\014");rup();provv_ant()}
      if(inpt==3){cat("\014");docoe();ai();dac();com_cig();ldo();reg_es();provv_liq()}
      if(inpt==3.1){cat("\014");docoe()}
      if(inpt==3.2){cat("\014");ai();com_cig()}
      if(inpt==3.3){cat("\014");dac()}
      if(inpt==3.4){cat("\014");ldo()}
      if(inpt==3.5){cat("\014");reg_es()}
      if(inpt==3.6){cat("\014");provv_liq()}
    }else{
      cat("

    Che documento vuoi generare?
    (per singoli documenti digitare la cifra in parentesi, ad es. 3.1 per Autocertificazioni operatore economico)
      1: RAS, Richiesta pagina web
      2: Nomina RUP, Provvedimento registrazione anticipata
      3: Autocertificazioni operatore economico (.1), Atto istruttorio e Comunicazione CIG (.2),
         Decisione a contrattare (.3), Lettera d'ordine (.4), Certificato di regolare esecuzione (.5),
         Provvedimento di liquidazione (.6)
      4: Assenza doppio finanziamento e Funzionalità del bene (.1), Checklist (.2)

")
      inpt <- readline()
      if(inpt==1){cat("\014");ras.pnrr();pag()}
      if(inpt==2){cat("\014");rup.pnrr();provv_ant()}
      if(inpt==3){cat("\014");docoe.pnrr();ai.pnrr();dac.pnrr();com_cig();ldo.pnrr();reg_es();provv_liq()}
      if(inpt==4){cat("\014");doppio_fin.pnrr();fun_bene.pnrr();chklst.pnrr()}
      if(inpt==3.1){cat("\014");docoe.pnrr()}
      if(inpt==3.2){cat("\014");ai.pnrr();com_cig()}
      if(inpt==3.3){cat("\014");dac.pnrr()}
      if(inpt==3.4){cat("\014");ldo.pnrr()}
      if(inpt==3.5){cat("\014");reg_es()}
      if(inpt==3.6){cat("\014");provv_liq()}
      if(inpt==4.1){cat("\014");doppio_fin.pnrr();fun_bene.pnrr()}
      if(inpt==4.2){cat("\014");chklst.pnrr()}
    }
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

