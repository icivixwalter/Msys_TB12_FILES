Attribute VB_Name = "MSys_TB12_Mdl01_TEST"
Option Compare Database
Option Explicit
'==================================================================================================
'  MODULO    : MSys_TB12_Mdl01_TEST
'  PROGETTO  : MSys_TB12_FILES  (mini-progetto TB12)
'  SCOPO     : test delle 4 classi MSys_TB12_Cls01_LOG, Cls02_FILE, Cls03_DATABASE, Cls04_OGGETTI.
'              I test automatici lavorano SOLO in %TEMP% (cartelle e database temporanei) e su oggetti
'              zz_TEST_* del database corrente, che vengono eliminati alla fine.
'  USO       : finestra Immediata ->  TEST_TUTTI            (tutti i test automatici + riepilogo)
'                                     TEST_Cls02_FILE       (una sola classe)
'                                     TEST_INTERATTIVI      (finestre di selezione: servono le risposte dell'utente)
'  LOG       : <cartella del database>\Logs\MSys_TB12_Mdl01_TEST_AAAA_MM_GG.txt  (stessa riga in Debug.Print)
'              ogni verifica: Pos (es. 04.02.02) | OK oppure KO | descrizione
'              le classi scrivono anche i propri log: alcune righe AVVISO/ERRORE sono VOLUTE (test dei casi di errore)
'  VERSIONE  : 1.0 - 2026-11-15
'--------------------------------------------------------------------------------------------------
'  INDICE DELLE SEZIONI, DEI PARAGRAFI E DEI SOTTOPARAGRAFI  (cerca l'etichetta, es. #04.02#)
'--------------------------------------------------------------------------------------------------
'  #01#  IMPOSTAZIONI = contatori e costanti
'  #02#  TEST_TUTTI = esegue i 4 test automatici e scrive il riepilogo
'  #03#  TEST_Cls01_LOG
'        #03.01# = scrittura INFO/ERRORI     #03.02# = errore con posizione
'        #03.03# = file personalizzato       #03.04# = manutenzione (conteggio, dimensione, pulizia)
'  #04#  TEST_Cls02_FILE
'        #04.01# = preparazione              #04.02# = percorso lungo oltre 300 caratteri
'        #04.03# = file nel percorso lungo   #04.04# = attributi, dimensione, date, parti del percorso
'        #04.05# = elenchi                   #04.06# = spostamento ed eliminazione
'        #04.07# = pulizia e funzioni di percorso
'  #05#  TEST_Cls03_DATABASE
'        #05.01# = database di prova         #05.02# = SQL, conteggi, valori, esistenza
'        #05.03# = query ed elenco di query  #05.04# = transazione annullata
'        #05.05# = collegamenti              #05.06# = compattazione e pulizia
'  #06#  TEST_Cls04_OGGETTI
'        #06.01# = elenco del database corrente  #06.02# = database di prova
'        #06.03# = importazione ed esportazione  #06.04# = accodamento di gruppo
'        #06.05# = inventario senza doppioni     #06.06# = pulizia
'  #07#  TEST_INTERATTIVI = finestre di dialogo, proprieta', apertura cartella lunga, visualizza log
'  #90#  SUPPORTO
'        #90.01# = VERIFICA     #90.02# = LOG_TEST     #90.03# = TEMP_Percorso     #90.04# = ELIMINA_SE_ESISTE
'==================================================================================================


'##################################################################################################
'## #01# IMPOSTAZIONI = contatori e costanti
'##################################################################################################
Private Const NOME_MODULO As String = "MSys_TB12_Mdl01_TEST"
Private m_Ok_l As Long
Private m_Ko_l As Long
Private m_ElencoKo_s As String


'##################################################################################################
'## #02# TEST_TUTTI = esegue i 4 test automatici e scrive il riepilogo
'##################################################################################################
Public Function TEST_TUTTI(Optional ByVal mostraMessaggio As Boolean = True) As String
    m_Ok_l = 0: m_Ko_l = 0: m_ElencoKo_s = vbNullString
    LOG_TEST "02.00", "------", "========== TEST DELLE CLASSI MSys_TB12 - " & Format$(Now, "dd/mm/yyyy hh:nn") & " =========="
    TEST_Cls01_LOG
    TEST_Cls02_FILE
    TEST_Cls03_DATABASE
    TEST_Cls04_OGGETTI
    TEST_TUTTI = "VERIFICHE: " & (m_Ok_l + m_Ko_l) & " - OK: " & m_Ok_l & " - KO: " & m_Ko_l
    LOG_TEST "02.99", IIf(m_Ko_l = 0, "FINE", "ERRORE"), TEST_TUTTI & IIf(m_Ko_l = 0, "", " -> " & m_ElencoKo_s)
    If mostraMessaggio Then MsgBox TEST_TUTTI & IIf(m_Ko_l = 0, "", vbCrLf & vbCrLf & "Non superate: " & m_ElencoKo_s) & vbCrLf & vbCrLf & _
                                   "Dettagli: Logs\" & NOME_MODULO & "_" & Format$(Date, "yyyy_mm_dd") & ".txt", _
                                   IIf(m_Ko_l = 0, vbInformation, vbExclamation), "Test classi MSys_TB12"
End Function


'##################################################################################################
'## #03# TEST_Cls01_LOG
'##################################################################################################
Public Sub TEST_Cls01_LOG()
Dim lg As MSys_TB12_Cls01_LOG, sPos_s As String, sCart_s As String
On Error GoTo ERRORE
    LOG_TEST "03.00", "INIZIO", "MSys_TB12_Cls01_LOG"
    Set lg = New MSys_TB12_Cls01_LOG
    sCart_s = TEMP_Percorso("MSys_TB12_TEST_LOG")
    If Len(Dir$(sCart_s, vbDirectory)) > 0 Then On Error Resume Next: Kill sCart_s & "\*.*": RmDir sCart_s: On Error GoTo ERRORE
    lg.CartellaLog = sCart_s
    lg.NomeBase = "TEST_LOG"
    lg.origine = NOME_MODULO
    lg.ScriviAVideo = False
    '//#03.01# = scrittura INFO ed ERRORI, creazione della cartella
    sPos_s = "03.01"
    lg.Separatore "prova"
    lg.ScriviInfo "TEST_Cls01_LOG", "messaggio informativo"
    lg.Scrivi "TEST_Cls01_LOG", "03.01.01", "OK", "riga generica"
    VERIFICA "03.01.01", "Cls01 crea la cartella dei log", Len(Dir$(sCart_s, vbDirectory)) > 0
    VERIFICA "03.01.02", "Cls01 file INFO del giorno", Len(Dir$(lg.PercorsoFile("INFO"))) > 0
    '//#03.02# = errore con posizione nel codice
    sPos_s = "03.02"
    lg.ScriviErrore "TEST_Cls01_LOG", "03.02.01", 999, "errore simulato dal test"
    VERIFICA "03.02.01", "Cls01 file ERRORI del giorno", Len(Dir$(lg.PercorsoFile("ERRORI"))) > 0
    VERIFICA "03.02.02", "Cls01 UltimoErrore riporta la posizione #03.02.01#", InStr(lg.UltimoErrore, "#03.02.01#") > 0
    VERIFICA "03.02.03", "Cls01 righe: 4 in INFO, 1 in ERRORI", lg.ContaRigheLog("TEST_LOG_INFO_*") = 4 And lg.ContaRigheLog("TEST_LOG_ERRORI_*") = 1
    '//#03.03# = file personalizzato
    sPos_s = "03.03"
    lg.ScriviLogPersonalizzato "TEST_Cls01_LOG", "riga personalizzata", "TEST_LOG_PERSONALIZZATO.txt"
    VERIFICA "03.03.01", "Cls01 file personalizzato", Len(Dir$(sCart_s & "\TEST_LOG_PERSONALIZZATO.txt")) > 0
    '//#03.04# = manutenzione
    sPos_s = "03.04"
    VERIFICA "03.04.01", "Cls01 DimensioneLog > 0", lg.DimensioneLog("TEST_LOG*") > 0
    lg.Attivo = False   'PulisciLog scrive una riga di esito: spento il log, la cartella resta vuota
    VERIFICA "03.04.02", "Cls01 PulisciLog elimina i 3 file di prova", lg.PulisciLog(-1, "TEST_LOG*") = 3
    VERIFICA "03.04.03", "Cls01 cartella vuota dopo la pulizia", Len(Dir$(sCart_s & "\*.*")) = 0
    RmDir sCart_s
USCITA:
    Set lg = Nothing
    LOG_TEST "03.99", "FINE", "MSys_TB12_Cls01_LOG"
    Exit Sub
ERRORE:
    VERIFICA sPos_s, "Cls01 errore imprevisto " & Err.Number & ": " & Err.Description, False
    Resume USCITA
End Sub


'##################################################################################################
'## #04# TEST_Cls02_FILE
'##################################################################################################
Public Sub TEST_Cls02_FILE()
Dim fs As MSys_TB12_Cls02_FILE, sPos_s As String, sBase_s As String, sLungo_s As String, sSeme_s As String
Dim sCopia_s As String, sSpost_s As String, n As Long, iF_i As Integer
On Error GoTo ERRORE
    LOG_TEST "04.00", "INIZIO", "MSys_TB12_Cls02_FILE"
    Set fs = New MSys_TB12_Cls02_FILE
    '//#04.01# = preparazione: cartella di prova pulita in %TEMP%
    sPos_s = "04.01"
    sBase_s = TEMP_Percorso("MSys_TB12_TEST_FILE")
    If fs.CartellaEsiste(sBase_s) Then fs.EliminaCartella sBase_s, True
    VERIFICA "04.01.01", "Cls02 CreaCartella di prova", fs.CreaCartella(sBase_s)
    '//#04.02# = percorso lungo oltre 300 caratteri, creato in un colpo solo (cartelle intermedie comprese)
    sPos_s = "04.02"
    sLungo_s = sBase_s
    Do While Len(sLungo_s) < 300
        n = n + 1
        sLungo_s = sLungo_s & "\LIVELLO_" & Format$(n, "00") & "_CARTELLA_DI_PROVA_PERCORSO_LUNGO"
    Loop
    VERIFICA "04.02.01", "Cls02 IsLongPath su " & Len(sLungo_s) & " caratteri", fs.IsLongPath(sLungo_s)
    VERIFICA "04.02.02", "Cls02 CreaCartella percorso lungo (" & n & " livelli)", fs.CreaCartella(sLungo_s)
    VERIFICA "04.02.03", "Cls02 CartellaEsiste percorso lungo", fs.CartellaEsiste(sLungo_s)
    VERIFICA "04.02.04", "Cls02 FileEsiste su una cartella = False", Not fs.FileEsiste(sLungo_s)
    '//#04.03# = file nel percorso lungo (il seme si scrive con VBA nel percorso corto)
    sPos_s = "04.03"
    sSeme_s = sBase_s & "\seme.txt"
    iF_i = FreeFile: Open sSeme_s For Output As #iF_i: Print #iF_i, "riga di prova MSys_TB12_Cls02_FILE": Close #iF_i
    sCopia_s = sLungo_s & "\copia_seme.txt"
    VERIFICA "04.03.01", "Cls02 FileEsiste percorso corto", fs.FileEsiste(sSeme_s)
    VERIFICA "04.03.02", "Cls02 CopiaFile nel percorso lungo", fs.CopiaFile(sSeme_s, sCopia_s)
    VERIFICA "04.03.03", "Cls02 FileEsiste percorso lungo (" & Len(sCopia_s) & " caratteri)", fs.FileEsiste(sCopia_s)
    VERIFICA "04.03.04", "Cls02 CopiaFile senza sovrascrivere un file esistente = False (riga ERRORE voluta)", Not fs.CopiaFile(sSeme_s, sCopia_s, False)
    '//#04.04# = attributi, dimensione, date, parti del percorso
    sPos_s = "04.04"
    VERIFICA "04.04.01", "Cls02 DimensioneFile = FileLen del seme", fs.DimensioneFile(sCopia_s) = FileLen(sSeme_s)
    VERIFICA "04.04.02", "Cls02 DataModifica di oggi", DateValue(fs.DataModifica(sCopia_s)) = Date
    VERIFICA "04.04.03", "Cls02 AttributiTesto contiene Archivio", InStr(fs.AttributiTesto(sCopia_s), "Archivio") > 0
    VERIFICA "04.04.04", "Cls02 ProprietaFile segnala il percorso lungo", InStr(fs.ProprietaFile(sCopia_s), "percorso lungo") > 0
    VERIFICA "04.04.05", "Cls02 GetFileName, GetFileExtension, GetFolderPath", _
             fs.GetFileName(sCopia_s) = "copia_seme.txt" And fs.GetFileExtension(sCopia_s) = "txt" And fs.GetFolderPath(sCopia_s) = sLungo_s
    VERIFICA "04.04.06", "Cls02 GetFileSizeFormatted in bytes", Right$(fs.GetFileSizeFormatted(sCopia_s), 5) = "bytes"
    '//#04.05# = elenchi
    sPos_s = "04.05"
    VERIFICA "04.05.01", "Cls02 ElencoFile *.txt nel percorso lungo = 1", fs.ElencoFile(sLungo_s, "*.txt").Count = 1
    VERIFICA "04.05.02", "Cls02 ElencoCartelle ricorsivo = " & n & " livelli", fs.ElencoCartelle(sBase_s, "*", True).Count = n
    VERIFICA "04.05.03", "Cls02 ElencoFile ricorsivo = 2", fs.ElencoFile(sBase_s, "*.txt", True).Count = 2
    '//#04.06# = spostamento ed eliminazione
    sPos_s = "04.06"
    sSpost_s = sLungo_s & "\rinominato.txt"
    VERIFICA "04.06.01", "Cls02 SpostaFile (rinomina) nel percorso lungo", fs.SpostaFile(sCopia_s, sSpost_s)
    VERIFICA "04.06.02", "Cls02 dopo lo spostamento l'origine non c'e' piu'", Not fs.FileEsiste(sCopia_s) And fs.FileEsiste(sSpost_s)
    SetAttr sSeme_s, vbReadOnly
    VERIFICA "04.06.03", "Cls02 EliminaFile su file in sola lettura", fs.EliminaFile(sSeme_s)
    VERIFICA "04.06.04", "Cls02 EliminaCartella non vuota senza forzare = False (riga ERRORE voluta)", Not fs.EliminaCartella(sBase_s, False)
    '//#04.07# = pulizia e funzioni di percorso
    sPos_s = "04.07"
    VERIFICA "04.07.01", "Cls02 EliminaCartella forzata con il percorso lungo dentro", fs.EliminaCartella(sBase_s, True)
    VERIFICA "04.07.02", "Cls02 cartella di prova eliminata", Not fs.CartellaEsiste(sBase_s)
    VERIFICA "04.07.03", "Cls02 NormalizzaPercorso locale e di rete", _
             fs.NormalizzaPercorso("C:\A\\B\") = "\\?\C:\A\B" And fs.NormalizzaPercorso("\\server\share\a") = "\\?\UNC\server\share\a"
    VERIFICA "04.07.04", "Cls02 PercorsoBreve di %ProgramFiles%", Len(fs.PercorsoBreve(Environ$("ProgramFiles"))) > 0
USCITA:
    Set fs = Nothing
    LOG_TEST "04.99", "FINE", "MSys_TB12_Cls02_FILE"
    Exit Sub
ERRORE:
    VERIFICA sPos_s, "Cls02 errore imprevisto " & Err.Number & ": " & Err.Description, False
    Resume USCITA
End Sub


'##################################################################################################
'## #05# TEST_Cls03_DATABASE
'##################################################################################################
Public Sub TEST_Cls03_DATABASE()
Dim dbc As MSys_TB12_Cls03_DATABASE, dbl As MSys_TB12_Cls03_DATABASE, sPos_s As String
Dim sDb_s As String, sDb2_s As String, sComp_s As String, el As Collection, x As DAO.database
On Error GoTo ERRORE
    LOG_TEST "05.00", "INIZIO", "MSys_TB12_Cls03_DATABASE"
    '//#05.01# = database di prova in %TEMP%
    sPos_s = "05.01"
    sDb_s = TEMP_Percorso("MSys_TB12_TEST_DB.mdb")
    sDb2_s = TEMP_Percorso("MSys_TB12_TEST_DB2.mdb")
    sComp_s = TEMP_Percorso("MSys_TB12_TEST_COMPATTATO.mdb")
    ELIMINA_SE_ESISTE sDb_s: ELIMINA_SE_ESISTE sDb2_s: ELIMINA_SE_ESISTE sComp_s: ELIMINA_SE_ESISTE sComp_s & ".src.mdb"
    On Error Resume Next: CurrentDb.TableDefs.Delete "zz_TEST_LINK_T_PROVA": On Error GoTo ERRORE
    Set x = DBEngine.CreateDatabase(sDb_s, dbLangGeneral, dbVersion40): x.Close: Set x = Nothing
    Set dbc = New MSys_TB12_Cls03_DATABASE
    VERIFICA "05.01.01", "Cls03 Apri database esterno", dbc.Apri(sDb_s)
    '//#05.02# = SQL, conteggi, valori, esistenza
    sPos_s = "05.02"
    VERIFICA "05.02.01", "Cls03 EseguiSQL CREATE TABLE", dbc.EseguiSQL("CREATE TABLE T_PROVA (ID COUNTER CONSTRAINT PK_T_PROVA PRIMARY KEY, NOME TEXT(50), VALORE LONG)")
    VERIFICA "05.02.02", "Cls03 EseguiSQL tre INSERT", dbc.EseguiSQL("INSERT INTO T_PROVA (NOME, VALORE) VALUES ('A', 1)") And _
             dbc.EseguiSQL("INSERT INTO T_PROVA (NOME, VALORE) VALUES ('B', 2)") And dbc.EseguiSQL("INSERT INTO T_PROVA (NOME, VALORE) VALUES ('C', 3)")
    VERIFICA "05.02.03", "Cls03 ContaRecord = 3", dbc.ContaRecord("T_PROVA") = 3
    VERIFICA "05.02.04", "Cls03 ContaRecord con criterio = 1", dbc.ContaRecord("T_PROVA", "VALORE > 2") = 1
    VERIFICA "05.02.05", "Cls03 ValoreSQL somma = 6", dbc.ValoreSQL("SELECT SUM(VALORE) FROM T_PROVA") = 6
    VERIFICA "05.02.06", "Cls03 TabellaEsiste e CampoEsiste", dbc.TabellaEsiste("T_PROVA") And dbc.CampoEsiste("T_PROVA", "NOME") And Not dbc.CampoEsiste("T_PROVA", "NON_ESISTE")
    VERIFICA "05.02.07", "Cls03 EseguiSQL errato = False (riga ERRORE voluta)", Not dbc.EseguiSQL("UPDATE TABELLA_INESISTENTE SET X = 1")
    '//#05.03# = query ed elenco di query (ex ClsManager)
    sPos_s = "05.03"
    dbc.DatabaseDAO.CreateQueryDef "Q_PROVA_AGGIORNA", "UPDATE T_PROVA SET VALORE = VALORE + 10"
    VERIFICA "05.03.01", "Cls03 QueryEsiste", dbc.QueryEsiste("Q_PROVA_AGGIORNA") And Not dbc.QueryEsiste("Q_NON_ESISTE")
    VERIFICA "05.03.02", "Cls03 EseguiQuery su 3 record", dbc.EseguiQuery("Q_PROVA_AGGIORNA") And dbc.RecordInteressati = 3
    dbc.AggiungiQuery el, "Q_PROVA_AGGIORNA", "aggiorna i valori"
    dbc.AggiungiQuery el, "Q_NON_ESISTE", "query mancante"
    VERIFICA "05.03.03", "Cls03 EseguiElencoQuery con una query mancante = False", Not dbc.EseguiElencoQuery(el)
    VERIFICA "05.03.04", "Cls03 UltimoReport: 1 eseguita", InStr(dbc.UltimoReport, "Eseguite: 1") > 0
    '//#05.04# = transazione annullata
    sPos_s = "05.04"
    dbc.IniziaTransazione
    dbc.EseguiSQL "DELETE FROM T_PROVA"
    dbc.AnnullaTransazione
    VERIFICA "05.04.01", "Cls03 AnnullaTransazione: i 3 record ci sono ancora", dbc.ContaRecord("T_PROVA") = 3
    dbc.Chiudi
    '//#05.05# = collegamenti nel database corrente
    sPos_s = "05.05"
    Set dbl = New MSys_TB12_Cls03_DATABASE
    dbl.Apri
    VERIFICA "05.05.01", "Cls03 CollegaTabella", dbl.CollegaTabella(sDb_s, "T_PROVA", "zz_TEST_LINK_T_PROVA")
    VERIFICA "05.05.02", "Cls03 tabella collegata leggibile (3 record)", dbl.ContaRecord("zz_TEST_LINK_T_PROVA") = 3
    '    secondo database con la stessa tabella VUOTA: dopo il ricollegamento il conteggio deve essere 0
    Set x = DBEngine.CreateDatabase(sDb2_s, dbLangGeneral, dbVersion40)
    x.Execute "CREATE TABLE T_PROVA (ID COUNTER, NOME TEXT(50), VALORE LONG)", dbFailOnError
    x.Close: Set x = Nothing
    VERIFICA "05.05.03", "Cls03 RicollegaTabelle verso il secondo database = 1", dbl.RicollegaTabelle(sDb2_s, "zz_TEST_LINK*") = 1
    VERIFICA "05.05.04", "Cls03 dopo il ricollegamento la tabella e' quella vuota (0 record)", dbl.ContaRecord("zz_TEST_LINK_T_PROVA") = 0
    VERIFICA "05.05.05", "Cls03 ElencoCollegamenti punta al secondo database", InStr(dbl.ElencoCollegamenti("zz_TEST_LINK*")(1), "DB2") > 0
    CurrentDb.TableDefs.Delete "zz_TEST_LINK_T_PROVA"
    dbl.Chiudi
    '//#05.06# = compattazione (su un database mai collegato, quindi libero) e pulizia
    sPos_s = "05.06"
    Set x = DBEngine.CreateDatabase(sComp_s & ".src.mdb", dbLangGeneral, dbVersion40): x.Close: Set x = Nothing
    VERIFICA "05.06.01", "Cls03 CompattaDatabase", dbc.CompattaDatabase(sComp_s & ".src.mdb", sComp_s) And Len(Dir$(sComp_s)) > 0
    ELIMINA_SE_ESISTE sDb_s: ELIMINA_SE_ESISTE sDb2_s: ELIMINA_SE_ESISTE sComp_s: ELIMINA_SE_ESISTE sComp_s & ".src.mdb"
USCITA:
    Set dbc = Nothing: Set dbl = Nothing
    LOG_TEST "05.99", "FINE", "MSys_TB12_Cls03_DATABASE"
    Exit Sub
ERRORE:
    VERIFICA sPos_s, "Cls03 errore imprevisto " & Err.Number & ": " & Err.Description, False
    On Error Resume Next
    CurrentDb.TableDefs.Delete "zz_TEST_LINK_T_PROVA"
    Resume USCITA
End Sub


'##################################################################################################
'## #06# TEST_Cls04_OGGETTI
'##################################################################################################
Public Sub TEST_Cls04_OGGETTI()
Dim og As MSys_TB12_Cls04_OGGETTI, sPos_s As String, sDb_s As String, x As DAO.database
On Error GoTo ERRORE
    LOG_TEST "06.00", "INIZIO", "MSys_TB12_Cls04_OGGETTI"
    Set og = New MSys_TB12_Cls04_OGGETTI
    '//#06.01# = elenco del database corrente
    sPos_s = "06.01"
    VERIFICA "06.01.01", "Cls04 le 4 classi MSys_TB12_Cls01..04 nel database corrente (escluso il separatore Cls00)", og.ContaOggetti("MODULI", "MSys_TB12_Cls0[1-4]_*") = 4
    VERIFICA "06.01.02", "Cls04 OggettoEsiste modulo di test", og.OggettoEsiste(NOME_MODULO, "MODULI")
    VERIFICA "06.01.03", "Cls04 TipoAccess FORMS = acForm e tipo sconosciuto = -1", og.TipoAccess("FORMS") = acForm And og.TipoAccess("XYZ") = -1
    '//#06.02# = database di prova con una tabella (2 righe) e una query
    sPos_s = "06.02"
    sDb_s = TEMP_Percorso("MSys_TB12_TEST_OGGETTI.mdb")
    ELIMINA_SE_ESISTE sDb_s
    '    oggetti rimasti da un'esecuzione interrotta
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "zz_TEST_Q_IMPORTATA": DoCmd.DeleteObject acTable, "zz_TEST_T_OGG": DoCmd.DeleteObject acTable, "zz_TEST_ELENCO"
    On Error GoTo ERRORE
    Set x = DBEngine.CreateDatabase(sDb_s, dbLangGeneral, dbVersion40)
    x.Execute "CREATE TABLE zz_TEST_T_OGG (ID COUNTER, NOME TEXT(50))", dbFailOnError
    x.Execute "INSERT INTO zz_TEST_T_OGG (NOME) VALUES ('uno')", dbFailOnError
    x.Execute "INSERT INTO zz_TEST_T_OGG (NOME) VALUES ('due')", dbFailOnError
    x.CreateQueryDef "zz_TEST_Q_OGG", "SELECT * FROM zz_TEST_T_OGG"
    x.Close: Set x = Nothing
    VERIFICA "06.02.01", "Cls04 ElencoOggetti esterno: 1 tabella e 1 query", og.ContaOggetti("TABLE", "*", sDb_s) = 1 And og.ContaOggetti("QUERY", "*", sDb_s) = 1
    '//#06.03# = importazione ed esportazione
    sPos_s = "06.03"
    VERIFICA "06.03.01", "Cls04 ImportaOggetto query con nuovo nome", og.ImportaOggetto(sDb_s, "QUERY", "zz_TEST_Q_OGG", "zz_TEST_Q_IMPORTATA")
    VERIFICA "06.03.02", "Cls04 la query importata esiste", og.OggettoEsiste("zz_TEST_Q_IMPORTATA", "QUERY")
    VERIFICA "06.03.03", "Cls04 seconda importazione senza sovrascrivere = False (riga AVVISO voluta)", Not og.ImportaOggetto(sDb_s, "QUERY", "zz_TEST_Q_OGG", "zz_TEST_Q_IMPORTATA")
    VERIFICA "06.03.04", "Cls04 EsportaOggetto nel database di prova", og.EsportaOggetto(sDb_s, "QUERY", "zz_TEST_Q_IMPORTATA", "zz_TEST_Q_ESPORTATA")
    VERIFICA "06.03.05", "Cls04 la query esportata esiste nel database di prova", og.OggettoEsiste("zz_TEST_Q_ESPORTATA", "QUERY", sDb_s)
    '//#06.04# = accodamento di gruppo: tabelle zz_TEST_T_* dal database di prova
    sPos_s = "06.04"
    VERIFICA "06.04.01", "Cls04 AccodaOggetti tabelle zz_TEST_T_* = 1", og.AccodaOggetti(sDb_s, "TABLE", "zz_TEST_T_*") = 1
    VERIFICA "06.04.02", "Cls04 tabella importata con i dati (2 righe)", DCount("*", "zz_TEST_T_OGG") = 2
    '//#06.05# = inventario senza doppioni in una tabella temporanea con i campi di TB03
    sPos_s = "06.05"
    CurrentDb.Execute "CREATE TABLE zz_TEST_ELENCO (COD_PROGETTO_s TEXT(45), TIPOGGETTO_s TEXT(50), NOMEOGGETTO_s TEXT(255), DATAINS_d DATETIME)", dbFailOnError
    og.TabellaElenco = "zz_TEST_ELENCO"
    VERIFICA "06.05.01", "Cls04 AccodaElencoInTabella: 4 classi aggiunte", og.AccodaElencoInTabella("TEST_TB12", vbNullString, "MODULI", "MSys_TB12_Cls0[1-4]_*") = 4
    VERIFICA "06.05.02", "Cls04 seconda volta: nessun doppione", og.AccodaElencoInTabella("TEST_TB12", vbNullString, "MODULI", "MSys_TB12_Cls0[1-4]_*") = 0
    og.TabellaElenco = "zz_TABELLA_INESISTENTE"
    VERIFICA "06.05.03", "Cls04 tabella di inventario mancante = -1 (riga ERRORE voluta)", og.AccodaElencoInTabella("TEST_TB12") = -1
USCITA:
    '//#06.06# = pulizia degli oggetti di prova
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "zz_TEST_Q_IMPORTATA"
    DoCmd.DeleteObject acTable, "zz_TEST_T_OGG"
    DoCmd.DeleteObject acTable, "zz_TEST_ELENCO"
    ELIMINA_SE_ESISTE sDb_s
    Set og = Nothing
    LOG_TEST "06.99", "FINE", "MSys_TB12_Cls04_OGGETTI"
    Exit Sub
ERRORE:
    VERIFICA sPos_s, "Cls04 errore imprevisto " & Err.Number & ": " & Err.Description, False
    Resume USCITA
End Sub


'##################################################################################################
'## #07# TEST_INTERATTIVI = servono le risposte dell'utente (da lanciare a mano)
'##################################################################################################
Public Sub TEST_INTERATTIVI()
Dim fs As New MSys_TB12_Cls02_FILE, lg As New MSys_TB12_Cls01_LOG, s As String, c As Collection, sLungo_s As String, n As Long
    '//#07.01# = selezione di un file e scheda proprieta'
    s = fs.SelezionaFile("Database Access (*.mdb;*.accdb)", "*.mdb;*.accdb", CurrentProject.path, "TEST - scegli un database")
    If Len(s) > 0 Then fs.MostraProprietaFile s
    '//#07.02# = selezione multipla
    Set c = fs.SelezionaFileMultipli(, , CurrentProject.path, "TEST - scegli piu' file")
    MsgBox "File scelti: " & c.Count, vbInformation, NOME_MODULO
    '//#07.03# = selezione di una cartella e salvataggio
    s = fs.SelezionaCartella(CurrentProject.path, "TEST - scegli una cartella")
    MsgBox "Cartella scelta: " & s, vbInformation, NOME_MODULO
    s = fs.SalvaFile(, , CurrentProject.path, "PROVA_SALVATAGGIO.txt", "TEST - salva con nome (non viene scritto nulla)")
    MsgBox "Percorso scelto: " & s, vbInformation, NOME_MODULO
    '//#07.04# = apertura in Esplora file di una cartella oltre 260 caratteri
    sLungo_s = TEMP_Percorso("MSys_TB12_TEST_APRI")
    Do While Len(sLungo_s) < 280
        n = n + 1: sLungo_s = sLungo_s & "\LIVELLO_" & Format$(n, "00") & "_CARTELLA_DI_PROVA_PERCORSO_LUNGO"
    Loop
    fs.CreaCartella sLungo_s
    If fs.ApriCartella(sLungo_s) Then MsgBox "Si e' aperta una cartella di " & Len(sLungo_s) & " caratteri." & vbCrLf & "Chiudi Esplora file e premi OK: la cartella di prova verra' eliminata.", vbInformation, NOME_MODULO
    fs.EliminaCartella TEMP_Percorso("MSys_TB12_TEST_APRI"), True
    '//#07.05# = log di oggi della classe di test
    lg.NomeBase = NOME_MODULO
    lg.ScriviInfo "TEST_INTERATTIVI", "test interattivi conclusi"
    lg.VisualizzaLog "INFO"
End Sub


'##################################################################################################
'## #90# SUPPORTO
'##################################################################################################
Private Sub VERIFICA(ByVal sPos As String, ByVal sDescr As String, ByVal bEsito As Boolean)
'//#90.01# = conta e scrive l'esito di una verifica
    If bEsito Then
        m_Ok_l = m_Ok_l + 1
        LOG_TEST sPos, "OK", sDescr
    Else
        m_Ko_l = m_Ko_l + 1
        m_ElencoKo_s = m_ElencoKo_s & " #" & sPos & "#"
        LOG_TEST sPos, "KO", sDescr
    End If
End Sub

Private Sub LOG_TEST(ByVal sPos As String, ByVal sEsito As String, ByVal sTesto As String)
'//#90.02# = stessa riga in Debug.Print e in <db>\Logs\MSys_TB12_Mdl01_TEST_AAAA_MM_GG.txt
Dim sRiga_s As String, sCart_s As String, iF_i As Integer
    On Error Resume Next
    sRiga_s = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Modulo: " & NOME_MODULO & " | Pos: " & Left$(sPos & Space$(8), 8) & _
              " | Esito: " & Left$(sEsito & Space$(6), 6) & " | " & sTesto
    Debug.Print sRiga_s
    sCart_s = CurrentProject.path & "\Logs"
    If Len(Dir$(sCart_s, vbDirectory)) = 0 Then MkDir sCart_s
    iF_i = FreeFile
    Open sCart_s & "\" & NOME_MODULO & "_" & Format$(Date, "yyyy_mm_dd") & ".txt" For Append As #iF_i
    Print #iF_i, sRiga_s
    Close #iF_i
End Sub

Private Function TEMP_Percorso(ByVal sNome As String) As String
'//#90.03# = percorso nella cartella temporanea di Windows
    TEMP_Percorso = Environ$("TEMP") & "\" & sNome
End Function

Private Sub ELIMINA_SE_ESISTE(ByVal sFile As String)
'//#90.04#
    On Error Resume Next
    If Len(Dir$(sFile)) > 0 Then SetAttr sFile, vbNormal: Kill sFile
End Sub

