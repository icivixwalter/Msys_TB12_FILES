# MSys_TB12_Cls04_OGGETTI

Oggetti di Access tra database: elenco di tabelle, query, form, report, macro e moduli del database corrente o di uno esterno; esistenza; importazione ed esportazione di singoli oggetti o di gruppi; registrazione dell'elenco in una tabella di inventario, senza doppioni. Sostituisce il vecchio `clsDBInspector`.

File: `OBJECT\CLASSI\MSys_TB12_Cls04_OGGETTI.cls` · Dipendenze: riferimento DAO · Esempi: `ESEMPI\MSys_TB12_Cls04_OGGETTI_ESEMPI.txt`

## Tipi di oggetto

Gli stessi di `MSys_DF02_TIPO_OGGETTO`: `TABLE`, `QUERY`, `FORMS`, `REPORTS`, `MACRO`, `MODULI`, più `TUTTI`. `TipoAccess` li converte nelle costanti di Access (`acTable` … `acModule`).

## Metodi

| Metodo | Restituisce | Cosa fa | Sezione |
| --- | --- | --- | --- |
| `ElencoOggetti([tipo], [filtroNome], [percorsoDb], [escludiSistema])` | Collection | voci `TIPO\|NOME`; senza percorso il database corrente | #03.01# |
| `ContaOggetti([tipo], [filtroNome], [percorsoDb])` | Long | numero di oggetti | #03.02# |
| `OggettoEsiste(nome, tipo, [percorsoDb])` | Boolean | nome esatto | #04.01# |
| `ImportaOggetto(sorgente, tipo, nome, [nomeDest], [sovrascrivi])` | Boolean | da un altro database a quello corrente (tabelle con i dati) | #05.01# |
| `EsportaOggetto(destinazione, tipo, nome, [nomeDest])` | Boolean | dal database corrente a un altro | #05.02# |
| `AccodaOggetti(sorgente, [tipo], [filtroNome], [sovrascrivi])` | Long | importa un gruppo; quanti importati | #05.03# |
| `AccodaElencoInTabella(codiceProgetto, [percorsoDb], [tipo], [filtroNome])` | Long | righe aggiunte all'inventario; -1 se la tabella manca | #06.01# |
| `AccodaTuttiGliOggetti(codiceProgetto, [mostraReport])` | Boolean | compatibile: tabelle, query e form del database corrente | #06.02# |
| `TipoAccess(tipo)`, `TipoValido(tipo)` | Long, Boolean | conversione dei tipi | #07# |

## Inventario

Tabella e campi predefiniti sono quelli del mini-progetto TB03, modificabili con le proprietà:

| Proprietà | Predefinito |
| --- | --- |
| `TabellaElenco` | `Msys_TB03_ELENCO_ObjProgetti` |
| `CampoProgetto` | `COD_PROGETTO_s` |
| `CampoTipo` | `TIPOGGETTO_s` |
| `CampoNome` | `NOMEOGGETTO_s` |
| `CampoData` | `DATAINS_d` (vuoto = non scritto) |

La tabella può essere locale o collegata. Una riga già presente (stesso progetto, tipo e nome) non viene ripetuta.

## Log

`<cartella del database>\Logs\MSys_TB12_Cls04_OGGETTI_AAAA_MM_GG.txt`. Ogni importazione o esportazione è scritta con l'oggetto e il database; un oggetto già presente e non sovrascritto produce un `AVVISO`, non un errore.

## Indice delle sezioni

- `#01#` IMPOSTAZIONI · `#02#` CICLO_DI_VITA_E_PROPRIETA
- `#03#` ELENCO: `#03.01#` (`.01` apertura, `.02` tabelle e query, `.03` contenitori, `.04` chiusura), `#03.02#`
- `#04#` ESISTENZA · `#05#` TRASFERIMENTO: `#05.01#` (`.01` controlli, `.02` esistente, `.03` importazione), `#05.02#` (`.01`, `.02`), `#05.03#` (`.01` elenco, `.02` importazione, `.03` esito)
- `#06#` INVENTARIO: `#06.01#` (`.01` tabella e campi, `.02` elenco, `.03` inserimento), `#06.02#`
- `#07#` CONVERSIONI · `#90#` LOG_INTERNO
