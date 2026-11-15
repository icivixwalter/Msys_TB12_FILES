# MSys_TB12_Cls02_FILE

File e cartelle con le **API Windows Unicode** a 32 e 64 bit. Gestisce percorsi fino a 32.767 caratteri con il prefisso `\\?\`, attributi, dimensioni, date, elenchi, apertura con l'applicazione predefinita e finestre di selezione.

File: `OBJECT\CLASSI\MSys_TB12_Cls02_FILE.cls` · Dipendenze: nessuna · Esempi: `ESEMPI\MSys_TB12_Cls02_FILE_ESEMPI.txt`

## Percorsi oltre 260 caratteri

Le API classiche di Windows e `FileSystemObject` si fermano a 260 caratteri (`MAX_PATH`). Questa classe:

- trasforma ogni percorso assoluto in `\\?\C:\...` (o `\\?\UNC\server\...` per la rete) con `NormalizzaPercorso`;
- chiama solo funzioni Unicode `...W` passando le stringhe con `StrPtr`;
- dichiara le API due volte con la compilazione condizionale: `#If VBA7` (`PtrSafe`, `LongPtr`, Office 2010 e successivi a 32 o 64 bit) e `#Else` (Office 2007 e precedenti);
- per aprire con la shell oltre 260 caratteri usa il nome breve 8.3 (`PercorsoBreve`).

Provato: cartella di 317 caratteri creata in un colpo solo, file di 332 caratteri copiato, letto, spostato ed eliminato.

## Metodi

| Metodo | Restituisce | Cosa fa | Sezione |
| --- | --- | --- | --- |
| `FileEsiste(percorso)` | Boolean | il file esiste (non una cartella) | #03.01# |
| `CartellaEsiste(percorso)` | Boolean | la cartella esiste | #03.02# |
| `IsLongPath(percorso)` | Boolean | oltre 260 caratteri | #03.03# |
| `NormalizzaPercorso(percorso)` | String | percorso con `\\?\` | #04.01# |
| `PercorsoSenzaPrefisso(percorso)` | String | toglie `\\?\` | #04.02# |
| `PercorsoBreve(percorso)` | String | nome breve 8.3 | #04.03# |
| `GetFileName`, `GetFolderPath`, `GetFileExtension`, `UnisciPercorso` | String | parti del percorso (solo testo) | #04.04# |
| `CreaCartella(percorso)` | Boolean | crea anche le cartelle intermedie | #05.01# |
| `EliminaCartella(percorso, [forzata])` | Boolean | con `forzata = True` anche il contenuto | #05.02# |
| `ElencoFile(cartella, [filtro], [ricorsivo])` | Collection | percorsi dei file | #05.03# |
| `ElencoCartelle(cartella, [filtro], [ricorsivo])` | Collection | percorsi delle sottocartelle | #05.04# |
| `CopiaFile(origine, destinazione, [sovrascrivi])` | Boolean | destinazione file o cartella; crea le cartelle mancanti | #06.01# |
| `SpostaFile(origine, destinazione, [sovrascrivi])` | Boolean | sposta o rinomina, anche tra dischi | #06.02# |
| `EliminaFile(percorso)` | Boolean | anche se in sola lettura | #06.03# |
| `AttributiFile`, `AttributiTesto` | Long, String | attributi (-1 se inesistente) | #07.01# |
| `DimensioneFile`, `GetFileSizeFormatted` | Double, String | dimensione oltre 2 GB | #07.02# |
| `DataCreazione`, `DataModifica`, `DataUltimoAccesso` | Date | ora locale | #07.03# |
| `ProprietaFile(percorso)` | String | scheda proprietà in testo | #07.04# |
| `MostraProprietaFile(percorso)` | – | scheda in una finestra | #07.05# |
| `ApriFile(percorso)`, `ApriCartella(percorso)` | Boolean | applicazione predefinita / Esplora file | #08# |
| `SelezionaFile`, `SelezionaFileMultipli`, `SelezionaCartella`, `SalvaFile` | String / Collection | finestre di dialogo (`Application.FileDialog`) | #09# |

I nomi pubblici della versione precedente sono rimasti uguali: il codice che li usava funziona ancora.

## Proprietà

`CartellaLog`, `LogAttivo` (predefinito `True`), `LogDettagliato` (predefinito `False`: solo esiti ed errori), `UltimoErrore` (sola lettura).

## Log

`<cartella del database>\Logs\MSys_TB12_Cls02_FILE_AAAA_MM_GG.txt`. Per gli errori di Windows la riga riporta il codice, per esempio `errore Windows 80` (il file esiste già), `5` (accesso negato), `145` (cartella non vuota), `3` (percorso non trovato).

## Indice delle sezioni

- `#01#` IMPOSTAZIONI: `#01.01#` costanti, `#01.02#` strutture API, `#01.03#` API per VBA7, `#01.04#` API per VBA6, `#01.05#` variabili
- `#02#` CICLO_DI_VITA_E_PROPRIETA
- `#03#` VERIFICHE: `#03.01#` FileEsiste, `#03.02#` CartellaEsiste (`.01` vuoto, `.02` attributi, `.03` esito), `#03.03#` IsLongPath
- `#04#` PERCORSI: `#04.01#` normalizzazione (`.01` pulizia, `.02` UNC, `.03` unità), `#04.02#`, `#04.03#` percorso breve, `#04.04#` parti
- `#05#` CARTELLE: `#05.01#` crea (`.01` esiste, `.02` radice, `.03` livelli, `.04` esito), `#05.02#` elimina, `#05.03#` file, `#05.04#` cartelle
- `#06#` FILE: `#06.01#` copia, `#06.02#` sposta, `#06.03#` elimina (`.01` controlli, `.02` sola lettura, `.03` API)
- `#07#` ATTRIBUTI_E_PROPRIETA · `#08#` APERTURA · `#09#` FINESTRE_DI_DIALOGO
- `#90#` LOG_INTERNO · `#91#` SUPPORTO: `#91.01#` attributi, `#91.02#` date, `#91.03#` nome da ricerca, `#91.04#` dimensione, `#91.05#` scansione, `#91.06#` svuotamento ricorsivo
