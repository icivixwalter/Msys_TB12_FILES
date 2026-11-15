# MSys_TB12_Cls01_LOG

Logger per il codice applicativo: form, moduli e altre classi. Scrive errori con la posizione nel codice, messaggi informativi e righe su file personalizzati, e fa la manutenzione dei log. Ogni riga va **identica** in `Debug.Print` e nel file.

File: `OBJECT\CLASSI\MSys_TB12_Cls01_LOG.cls` · Dipendenze: nessuna · Esempi: `ESEMPI\MSys_TB12_Cls01_LOG_ESEMPI.txt`

## File di log

`<cartella del database>\Logs\<NomeBase>_<TIPO>_AAAA_MM_GG.txt`

- `NomeBase` predefinito: `MSys_TB12_Cls01_LOG`. Conviene impostarlo con il nome del modulo che usa il logger (es. `Form_MASTER`).
- `TIPO`: `ERRORI` (solo errori) oppure `INFO` (tutto il resto, più una copia degli errori, così la sequenza completa si legge in un solo file).
- Un file al giorno: `PulisciLog` elimina quelli vecchi.

Riga tipo:

```
2026-11-15 10:30:12 | Origine: Form_MASTER | Funzione: Form_Load | Pos: 03.02    | Esito: OK     | pagine caricate
```

## Metodi

| Metodo | Parametri | Cosa fa | Sezione |
| --- | --- | --- | --- |
| `Scrivi` | funzione, posizione, esito, testo | riga generica; `ERRORE` va nel file ERRORI | #03.01# |
| `ScriviErrore` | funzione, posizione, numero, descrizione | errore con la posizione nel codice; aggiorna `UltimoErrore` | #03.02# |
| `ScriviLogErrore` | nomeFunzione, numero, descrizione, [riga] | compatibile con la versione precedente | #03.03# |
| `ScriviInfo` / `ScriviLogEvento` | nomeFunzione, messaggio | messaggio informativo (compatibili) | #03.04# |
| `ScriviLogPersonalizzato` | nomeFunzione, messaggio, nomeFile | riga in un file scelto dentro `Logs` | #03.05# |
| `Separatore` | [titolo] | riga di separazione | #03.06# |
| `CartellaLogs` | – | percorso della cartella (creata se manca) | #04.01# |
| `PercorsoFile` | [tipo], [giorno] | percorso del file del giorno | #04.02# |
| `PulisciLog` | [giorni = 30], [filtro = *.txt] | elimina i file più vecchi; restituisce quanti | #05.01# |
| `DimensioneLog` | [filtro] | byte occupati | #05.02# |
| `ContaRigheLog` (`ContaLinieLog`) | [filtro] | righe totali | #05.03# |
| `VisualizzaLog` | [tipo = ERRORI] | apre il log di oggi | #05.04# |

## Proprietà

| Proprietà | Predefinito | Uso |
| --- | --- | --- |
| `CartellaLog` | `<db>\Logs` | cartella diversa per i log |
| `NomeBase` | `MSys_TB12_Cls01_LOG` | prima parte del nome dei file |
| `Origine` | nome della classe | testo della colonna Origine |
| `Attivo` | `True` | `False` = non scrive nulla |
| `ScriviAVideo` | `True` | `False` = niente `Debug.Print` |
| `UltimoErrore` | – | sola lettura: ultimo errore con posizione |

## Esempio rapido

```
Dim lg As New MSys_TB12_Cls01_LOG
lg.NomeBase = "Form_MASTER": lg.Origine = "Form_MASTER"
lg.Scrivi "Form_Load", "03.01", "INIZIO", "apertura form"
...
ERRORE:
    lg.ScriviErrore "Form_Load", sPos_s, Err.Number, Err.Description
```

## Indice delle sezioni

- `#01#` IMPOSTAZIONI: `#01.01#` costanti, `#01.02#` variabili
- `#02#` CICLO_DI_VITA_E_PROPRIETA: `#02.01#` Initialize, `#02.02#` Terminate, `#02.03#` proprietà
- `#03#` SCRITTURA: `#03.01#` Scrivi, `#03.02#` ScriviErrore (`.01` ERRORI, `.02` copia in INFO), `#03.03#` ScriviLogErrore, `#03.04#` ScriviInfo/Evento, `#03.05#` personalizzato, `#03.06#` separatore
- `#04#` FILE_DI_LOG: `#04.01#` cartella (`.01` scelta, `.02` creazione), `#04.02#` percorso, `#04.03#` scrittura (`.01` video, `.02` file, `.03` emergenza)
- `#05#` MANUTENZIONE: `#05.01#` pulizia (`.01` scadenza, `.02` eliminazione), `#05.02#` dimensione, `#05.03#` righe, `#05.04#` visualizza
- `#06#` EMERGENZA: `#06.01#` file in `%TEMP%` se la cartella non è scrivibile

## Rispetto alla versione precedente

- Un file al giorno e per tipo, con nome che indica chi l'ha scritto.
- Posizione nel codice in ogni riga.
- Nessun messaggio a video in caso di errore del logger: il programma non si ferma.
