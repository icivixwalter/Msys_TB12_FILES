# MSys_TB12 – Master Plan delle classi

Mini-progetto **TB12_FILES** del super-progetto MSys: quattro classi indipendenti per gestire file e cartelle, log, database e oggetti di Access. Ogni classe si importa da sola in qualsiasi progetto Access (32 o 64 bit) e scrive il proprio log nella sottocartella `Logs` del database.

Aggiornato il 15/11/2026. Test automatici: **68 verifiche, 68 superate**.

## Le quattro classi

| Classe | Cosa fa | Sostituisce |
| --- | --- | --- |
| `MSys_TB12_Cls01_LOG` | log per il codice applicativo: errori con la posizione nel codice, messaggi, file personalizzati, pulizia | `Msys_TB12_FILES_CLASSE_clsLogger`, `CLASSE_clsLogger` |
| `MSys_TB12_Cls02_FILE` | file e cartelle con API Windows Unicode a 32 e 64 bit, percorsi oltre 260 caratteri, attributi, date, elenchi, apertura, finestre di selezione | `Msys_TB12_FILES_CLASSE_clsFileSystem`, `CLASSE_clsFileSystem` |
| `MSys_TB12_Cls03_DATABASE` | database corrente o esterno: recordset, SQL, query singole o in elenco, transazioni, tabelle collegate, compattazione | `Msys_TB12_FILES_CLASSE_clsDatabase`, `Msys_TB12_FILES_CLASSE_ClsManager`, `CLASSE_ClsManager` |
| `MSys_TB12_Cls04_OGGETTI` | oggetti di Access tra database: elenco, esistenza, importazione, esportazione, accodamento di gruppo, inventario in tabella | `Msys_TB12_FILES_CLASSE_clsDBInspector`, `CLASSE_CLASSE_clsDBInspector` |
| `MSys_TB12_Mdl01_TEST` | modulo di test unico per le quattro classi | `CLSModulo_GESTIONE_OGGETTI_TEST_TutteLeClassi`, `TEST_clsFileSystem` |

Le versioni precedenti sono state tolte dal database e salvate in testo in `OBJECT\CLASSI_PRECEDENTI_2026_11_14_23_26_43\`.

## Come si collegano

```
                 codice dell'applicazione (form, moduli)
                    |          |            |            |
                    v          v            v            v
          Cls01_LOG   Cls02_FILE   Cls03_DATABASE   Cls04_OGGETTI      <- nessuna dipende dalle altre
                    \          |            |            /
                     \         |            |           /
                      v        v            v          v
                 <cartella del database>\Logs\
                   MSys_TB12_Cls01_LOG_ERRORI_AAAA_MM_GG.txt  (o <NomeBase>_ERRORI_...)
                   MSys_TB12_Cls02_FILE_AAAA_MM_GG.txt
                   MSys_TB12_Cls03_DATABASE_AAAA_MM_GG.txt
                   MSys_TB12_Cls04_OGGETTI_AAAA_MM_GG.txt
                   MSys_TB12_Mdl01_TEST_AAAA_MM_GG.txt
```

Le classi sono **indipendenti**: ognuna ha il proprio log interno, quindi in un altro progetto si può importare anche una sola classe. La `Cls01_LOG` serve al codice dell'applicazione, non alle altre classi.

## Regole comuni a tutte le classi

- **Nomi**: `MSys_TB12_Cls##_DESCRIZIONE`, secondo la grammatica del Master Plan MSys (progetto, mini-progetto, tipo e numero, descrizione).
- **Indice in testata**: ogni classe elenca sezioni `#SS#`, paragrafi `#SS.PP#` e sottoparagrafi `#SS.PP.ss#` con una descrizione breve. Nel codice ogni punto ha la stessa etichetta: si trova con Ctrl+F.
- **Log**: una riga per evento, identica in `Debug.Print` e nel file:
  `2026-11-15 10:30:12 | Classe: MSys_TB12_Cls02_FILE | Funzione: CopiaFile | Pos: 06.01.03 | Esito: ERRORE | copia non riuscita (errore Windows 80): ...`
- **Esiti**: `INIZIO`, `OK`, `INFO`, `AVVISO`, `ERRORE`, `FINE`.
- **Posizione dell'errore**: ogni procedura aggiorna la posizione corrente; in caso di errore il log scrive `codice fermo a #SS.PP.ss#` e la proprietà `UltimoErrore` la riporta al chiamante.
- **Livello di dettaglio**: per impostazione si scrivono gli esiti delle operazioni e tutti gli errori. Con `LogDettagliato = True` si scrive anche ogni passo intermedio.
- **Cartella Logs**: creata automaticamente se manca. Si può cambiare con la proprietà `CartellaLog`.

## Struttura della cartella

```
Msys_TB12_FILES\
├─ 00_DOC_MASTER_PLAN\        questo documento e una scheda per classe (.md e .html)
│  ├─ ESEMPI\                 un file .txt di esempi per ogni classe
│  └─ STRUMENTI\              GENERA_HTML.bat + md2html.py: rigenera gli .html dai .md
├─ MDB\
│  ├─ Msys_TB12_FILES.mdb     classi, modulo di test, form
│  ├─ Logs\                   log delle classi e dei test
│  └─ BACKUP_CLAUDE_...\      copia del database prima della modifica
└─ OBJECT\
   ├─ CLASSI\                 le 4 classi e il test in .cls/.bas, pronti da importare
   └─ CLASSI_PRECEDENTI_...\  le classi sostituite, in testo
```

## Difetti corretti rispetto alle classi precedenti

1. **Percorsi lunghi non funzionanti**: la classe FileSystem usava le API ANSI (`...A`), che non accettano il prefisso `\\?\`. Ora tutte le API sono Unicode (`...W`).
2. **Scrittura fuori memoria**: `GetFileAttributesEx` riceveva una variabile di 4 byte al posto di una struttura di 36. Ora riceve la struttura `WIN32_FILE_ATTRIBUTE_DATA` completa.
3. **Copia, spostamento, eliminazione e creazione oltre 260 caratteri**: passavano da `FileSystemObject`, che non li supporta. Ora usano `CopyFileW`, `MoveFileExW`, `DeleteFileW`, `CreateDirectoryW`, `RemoveDirectoryW`.
4. **Finestra di selezione a 64 bit**: la struttura `OPENFILENAME` aveva `lCustData` come `Long` e la dimensione calcolata con `Len`. Ora le finestre usano `Application.FileDialog`, che funziona a 32 e 64 bit, con selezione multipla e scelta della cartella.
5. **CreaCartella**: non creava le cartelle intermedie. Ora le crea tutte.
6. **DBInspector**: eseguiva tre query di accodamento che nel database non esistono. Ora elenca gli oggetti e li scrive direttamente nell'inventario, senza doppioni.
7. **Manager**: aveva scritte nel codice 15 query di un altro progetto (PRES3000, LLPP_ATTI). Ora l'elenco si passa a `EseguiElencoQuery`.
8. **Logger**: in caso di errore mostrava un messaggio che bloccava il programma. Ora scrive un file di emergenza in `%TEMP%` e non si ferma.

## Test

- **Automatici**: nella finestra Immediata `TEST_TUTTI` (oppure `TEST_Cls02_FILE` per una sola classe). Lavorano in `%TEMP%` e su oggetti `zz_TEST_*`, che vengono eliminati alla fine. Il test della classe FILE crea davvero un percorso di 332 caratteri e ci copia, sposta ed elimina file.
- **Interattivi**: `TEST_INTERATTIVI` apre le finestre di selezione, la scheda proprietà e una cartella di oltre 280 caratteri in Esplora file.
- **Risultato**: `Logs\MSys_TB12_Mdl01_TEST_AAAA_MM_GG.txt`, una riga per verifica con la posizione (es. `04.02.02`) e OK o KO. Nei log delle classi alcune righe `ERRORE` e `AVVISO` sono volute: sono i test dei casi di errore.

## Come usare le classi in un altro progetto

1. Nell'editor VBA del progetto di destinazione: **File › Importa file…** e scegli i `.cls` in `OBJECT\CLASSI\`.
2. Per `Cls03_DATABASE` e `Cls04_OGGETTI` serve il riferimento DAO (**Strumenti › Riferimenti › Microsoft Office … Access database engine Object Library**), presente di norma in ogni database Access.
3. **Debug › Compila**. Gli esempi pronti sono in `00_DOC_MASTER_PLAN\ESEMPI\`.

## Da decidere

- La form `fileDialog` contiene codice di un altro progetto (Codice Opera LLPP) con proprie API e propri log. Si può collegare a `MSys_TB12_Cls02_FILE` oppure spostarla nel progetto LLPP.
- La form `Msys_TB12_FILES_FrmM01_GESTIONI` segue ancora il vecchio schema di nome. Secondo la grammatica diventerebbe `MSys_TB12_Frm01_M01_GESTIONI`.
- `AccodaTuttiGliOggetti` scrive in `Msys_TB03_ELENCO_ObjProgetti`: in questo database la tabella non è collegata. Va collegata, oppure si indica un'altra tabella con la proprietà `TabellaElenco`.
