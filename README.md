# REPORT TOP30 AUTOMATIZZATO IN R
Script per la generazione automatica di report Excel e PDF a partire da file  esportati dal gestionale aziendale Junak4. Il flusso include: pulizia dei dati, calcolo di metriche (MLVE), aggiornamento di un template Excel preformattato e esportazione in PDF tramite automazione. Pensato per semplificare la reportistica e ridurre gli errori manuali.

## Descrizione

Script in R per la generazione automatica di report Excel e PDF a partire da file `.xlsx` esportati dal gestionale aziendale **Junak4**.  
Il flusso include:

- Pulizia e trasformazione dei dati
- Calcolo del margine lordo su venduto equivalente (MLVE)
- Aggiornamento di un template Excel preformattato
- Esportazione finale in PDF tramite automazione COM

Pensato per semplificare la reportistica periodica e ridurre gli errori manuali.

---

![ANTEPRIMA_REPORT](https://github.com/carchedimarco88-jpg/REPORT-TOP30/raw/main/MIGLIORI%2030%20PRODOTTI%20.png)

---


### Requisiti

- Sistema operativo: **Windows** (necessario per `RDCOMClient`)
- R ≥ 4.0
- Microsoft Excel installato
- Pacchetti R richiesti:
  ```r
  install.packages(c("openxlsx", "readxl", "tidyverse", "scales"))
  devtools::install_github("omegahat/RDCOMClient")  # solo su Windows

  
