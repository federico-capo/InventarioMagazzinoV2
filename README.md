# InventarioMagazzinoV2

## Descrizione
**InventarioMagazzinoV2** è un software progettato per aggiornare un file Excel di inventario utilizzando i dati estratti da un altro file Excel. Attraverso una semplice interfaccia, l'utente può importare i dati, aggiornare un secondo file e generare una copia aggiornata.

## Funzionalità principali

### 1. Lettura del file di configurazione INI
Il software legge un file `config.ini` per ottenere il percorso del secondo file Excel da aggiornare. Assicurarsi che il percorso nel file INI sia corretto.

### 2. Importazione dell'inventario
L'utente seleziona un file Excel contenente l'inventario. Il software estrae automaticamente i codici articolo, i nomi e le quantità dalla **riga 2** e dalle **colonne A, B e C**. 

- **Colonna A**: Codici articolo
- **Colonna B**: Nomi degli articoli
- **Colonna C**: Quantità degli articoli

### 3. Aggiornamento del secondo file Excel
Il software aggiorna il secondo file Excel con i dati estratti dall'inventario. Se un codice articolo non è presente nel file di destinazione, verrà aggiunto automaticamente.

### 4. Salvataggio del file aggiornato
Il file Excel aggiornato viene salvato con un suffisso `_agg.xlsx`. Verifica sempre il file aggiornato nella posizione indicata dal programma.

## Cosa deve fare l'utente

1. **Configurare il percorso del file Excel nel `config.ini`**.
2. **Selezionare il file Excel dell'inventario**.
3. **Verificare il file aggiornato** nella posizione indicata.
4. In caso di problemi nell'aggiornamento del file, **verificare il formato del file** e il percorso nel `config.ini`.

## Librerie necessarie

Per eseguire il software, è necessario avere installato le seguenti librerie:

- **EPPlus**: per la gestione dei file Excel.
- **IniParser**: per la lettura del file di configurazione INI.

## Installazione delle librerie

### Utilizzando Visual Studio

Apri il terminale o il **Package Manager Console** e digita:

```bash
Install-Package EPPlus
Install-Package IniParser
