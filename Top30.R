# Installazione pacchetti 
# install.packages("openxlsx")
#install.packages("devtools")
#devtools::install_github("omegahat/RDCOMClient")

# ======================== Parametri modificabili ==============================
data_riferimento <- "23 settembre 2025"

# Percorso file template Excel
template_file <- "PERCORSO_TEMPLATE.xlsx"

# Nomi file di output con data
output_file <- sprintf("TOP30 ARTICOLI %s.xlsx", data_riferimento)
pdf_file    <- sprintf("TOP30 ARTICOLI %s.pdf", data_riferimento)

# Percorso PDF sul Desktop
pdf_path <- file.path(Sys.getenv("USERPROFILE"), "Desktop", pdf_file)

# ================================= Librerie ===================================
library(RDCOMClient)
library(openxlsx)
library(tidyverse)
library(readxl)
library(scales)

message("=== Avvio script: ", Sys.time(), " ===")

# =========================== Import dati ================================
# Usa file.choose() per selezionare il file Excel di input
TOP30 <- read_excel(file.choose())

# --- Asserzioni minime ---
stopifnot(nrow(TOP30) > 0)  # Deve esserci almeno una riga
colonne <- c("kg1_description", "kg2_description", "kquantity_dec", 
                    "kqstock_dec", "ksale_dec", "krevenue_dec")
stopifnot(all(colonne %in% names(TOP30)))  # Devono esserci tutte le colonne

message("Dati importati correttamente. Righe: ", nrow(TOP30))

# ========================= Pulizia e trasformazione ===========================
# - Seleziona solo le colonne utili
# - Rimuove la prima riga (probabile intestazione extra)
# - Rinomina le colonne in italiano
# - Converte i campi numerici
# - Calcola MLVE
# - Ordina per venduti
# - Separa il fornitore da eventuali note tra parentesi
# - Tiene solo i primi 30
# - Arrotonda Ricavo e Margine a 2 decimali
TOP30 <- TOP30 %>% 
  select(all_of(colonne)) %>% 
  slice(-1) %>% 
  rename(
    Fornitore = kg1_description,
    Descrizione_articolo = kg2_description, 
    Venduti = kquantity_dec, 
    Giacenza = kqstock_dec, 
    Ricavo = ksale_dec, 
    Margine = krevenue_dec
  ) %>% 
  mutate(across(c(Venduti, Giacenza, Ricavo, Margine), as.numeric),
         MLVE = Margine / Ricavo) %>% 
  arrange(desc(Venduti)) %>% 
  separate(Fornitore, into = c("Fornitore", "x"), sep = "\\(", extra = "drop") %>% 
  select(-x) %>% 
  slice_head(n = 30) 

# ============================= Riga totale ===============================
# Calcola i totali complessivi e li aggiunge in fondo alla tabella
RigaTotale <- TOP30 %>%
  summarise(
    Fornitore = "Totale complessivo",
    Descrizione_articolo = "",
    Venduti = sum(Venduti, na.rm = TRUE),
    Giacenza = sum(Giacenza, na.rm = TRUE),
    Ricavo = sum(Ricavo, na.rm = TRUE),
    Margine = sum(Margine, na.rm = TRUE),
    MLVE = sum(Margine, na.rm = TRUE) / sum(Ricavo, na.rm = TRUE)
  )

TOP30 <- bind_rows(TOP30, RigaTotale)

message("Trasformazione completata. Righe finali: ", nrow(TOP30))

# --- Controllo esistenza template ---
stopifnot(file.exists(template_file))

# ========================== Aggiornamento Excel =============================
# Carica il file template
wb <- loadWorkbook(template_file)

# Scrive la data di riferimento nella cella A4:H4
writeData(wb, sheet = 1, x = data_riferimento, startCol = 1, startRow = 4)

# Scrive la tabella a partire da B5
writeData(wb, sheet = 1, x = TOP30, startCol = 2, startRow = 5, colNames = TRUE)

# Salva il nuovo file Excel con data nel nome
saveWorkbook(wb, output_file, overwrite = TRUE)
message("File Excel aggiornato: ")

# ======================= Esporta PDF con Excel COM ==========================
# Avvia Excel in background
excel <- COMCreate("Excel.Application")
excel[["Visible"]] <- FALSE

# Apre il file Excel appena creato
wb_com <- excel[["Workbooks"]]$Open(normalizePath(output_file))

# Esporta in PDF sul Desktop
wb_com$ExportAsFixedFormat(
  Type = 0,  # PDF
  Filename = normalizePath(pdf_path),
  Quality = 0,
  IncludeDocProperties = TRUE,
  IgnorePrintAreas = FALSE,
  OpenAfterPublish = FALSE
)

# Chiude il file e Excel
wb_com$Close(FALSE)
excel$Quit()

message("PDF creato sul Desktop: ")
message("=== Script completato: ", Sys.time(), " ===")
