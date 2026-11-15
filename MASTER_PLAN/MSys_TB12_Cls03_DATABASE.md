# MSys_TB12_Cls03_DATABASE

Database corrente o esterno con DAO: recordset, SQL di comando, query salvate (singole o in elenco), controlli di esistenza, conteggi, transazioni, tabelle collegate e compattazione. Riunisce la vecchia `clsDatabase` e il vecchio `ClsManager`.

File: `OBJECT\CLASSI\MSys_TB12_Cls03_DATABASE.cls` · Dipendenze: riferimento DAO · Esempi: `ESEMPI\MSys_TB12_Cls03_DATABASE_ESEMPI.txt`

## Metodi

| Metodo | Restituisce | Cosa fa | Sezione |
| --- | --- | --- | --- |
| `Apri([percorso], [solaLettura])` | Boolean | senza percorso: database corrente; con percorso: esterno | #03.01# |
| `GetDatabase()` | DAO.Database | database in uso (compatibile) | #03.02# |
| `Chiudi()` | – | chiude l'esterno, rilascia il corrente | #03.03# |
| `ApriRecordset(sql, [db], [tipo])` | DAO.Recordset | Nothing in caso di errore | #04.01# |
| `ContaRecord(origine, [criterio])` | Long | -1 in caso di errore | #04.02# |
| `ValoreSQL(sql, [seVuoto])` | Variant | primo campo del primo record | #04.03# |
| `EseguiSQL(sql, [db])` | Boolean | `RecordInteressati` = righe toccate | #05.01# |
| `EseguiQuery(nome, [usaDoCmd])` | Boolean | query di comando salvata; `usaDoCmd` se legge valori dalle form | #05.02# |
| `QueryEsiste`, `TabellaEsiste`, `CampoEsiste` | Boolean | senza leggere MSysObjects | #06# |
| `EseguiElencoQuery(elenco, [mostraReport], [usaDoCmd])` | Boolean | elenco di `Array(nome, descrizione)`; report in `UltimoReport` | #07.01# |
| `AggiungiQuery(elenco, nome, [descrizione])` | – | costruisce l'elenco | #07.02# |
| `IniziaTransazione`, `ConfermaTransazione`, `AnnullaTransazione` | Boolean | transazione sul workspace predefinito | #08# |
| `CollegaTabella(percorsoDb, tabella, [nomeLocale])` | Boolean | crea o ricrea un collegamento | #09.01# |
| `RicollegaTabelle(nuovoPercorso, [filtroNome], [vecchioPercorso])` | Long | quanti collegamenti aggiornati | #09.02# |
| `ElencoCollegamenti([filtroNome])` | Collection | voci `nome\|percorso` | #09.03# |
| `CompattaDatabase(origine, destinazione)` | Boolean | database chiuso, destinazione nuova | #10.01# |

## Dal vecchio Manager a EseguiElencoQuery

Il `ClsManager` aveva 15 query di altri progetti scritte nel codice (PRES3000, LLPP_ATTI…). Ora la classe non conosce nessuna query: l'elenco lo prepara chi la usa.

```
Dim dbc As New MSys_TB12_Cls03_DATABASE, el As Collection
dbc.AggiungiQuery el, "PRES3000_Qry01_10_AGGIORNA_Campo_IDKEY_CALENDARIO", "Aggiorna ID Key Calendario"
dbc.AggiungiQuery el, "LLPP_ATTI_Qy01-11_AGGIORNA_ANNO", "Aggiorna Anno Atto"
If Not dbc.EseguiElencoQuery(el, True) Then Debug.Print dbc.UltimoReport
```

## Proprietà

`DatabaseDAO`, `Percorso`, `RecordInteressati`, `UltimoReport`, `CartellaLog`, `LogAttivo`, `LogDettagliato`, `UltimoErrore`.

## Log

`<cartella del database>\Logs\MSys_TB12_Cls03_DATABASE_AAAA_MM_GG.txt`. Ogni comando SQL eseguito è scritto con il numero di record toccati; ogni errore con il testo SQL e la posizione.

## Indice delle sezioni

- `#01#` IMPOSTAZIONI · `#02#` CICLO_DI_VITA_E_PROPRIETA
- `#03#` CONNESSIONE: `#03.01#` Apri (`.01` chiusura precedente, `.02` corrente, `.03` esterno), `#03.02#`, `#03.03#`
- `#04#` LETTURA · `#05#` SCRITTURA: `#05.01#` (`.01` database, `.02` Execute), `#05.02#` (`.01` esistenza, `.02` esecuzione, `.03` esito)
- `#06#` CONTROLLI · `#07#` ELENCO_QUERY: `#07.01#` (`.01` controlli, `.02` ciclo, `.03` report), `#07.02#`
- `#08#` TRANSAZIONI · `#09#` COLLEGAMENTI: `#09.01#` (`.01` esistente, `.02` creazione), `#09.02#` (`.01` filtro, `.02` RefreshLink), `#09.03#`
- `#10#` MANUTENZIONE · `#90#` LOG_INTERNO
