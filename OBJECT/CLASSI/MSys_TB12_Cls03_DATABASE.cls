VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSys_TB12_Cls03_DATABASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'==================================================================================================
'  CLASSE    : MSys_TB12_Cls03_DATABASE
'  PROGETTO  : MSys_TB12_FILES  (mini-progetto TB12)
'  SCOPO     : database corrente o esterno con DAO: recordset, SQL, query (singole o in elenco),
'              controlli di esistenza, conteggi, transazioni, tabelle collegate, compattazione.
'  SOSTITUISCE: Msys_TB12_FILES_CLASSE_clsDatabase  e  Msys_TB12_FILES_CLASSE_ClsManager
'              (il Manager eseguiva 15 query di un altro progetto scritte nel codice: ora l'elenco
'               si passa a EseguiElencoQuery, vedi il file di esempi)
'  DIPENDENZE: riferimento DAO (Microsoft Office Access database engine Object Library). Nessuna classe.
'  LOG       : <cartella del database>\Logs\MSys_TB12_Cls03_DATABASE_AAAA_MM_GG.txt
'  VERSIONE  : 1.0 - 2026-11-15
'--------------------------------------------------------------------------------------------------
'  INDICE DELLE SEZIONI, DEI PARAGRAFI E DEI SOTTOPARAGRAFI  (cerca l'etichetta, es. #07.01.02#)
'--------------------------------------------------------------------------------------------------
'  #01#  IMPOSTAZIONI = costanti e variabili
'        #01.01# = costanti          #01.02# = variabili della classe
'  #02#  CICLO_DI_VITA_E_PROPRIETA = inizializzazione, chiusura, proprieta
'        #02.01# = Class_Initialize  #02.02# = Class_Terminate (chiude il database esterno)
'        #02.03# = proprieta: DatabaseDAO, Percorso, RecordInteressati, UltimoReport, log
'  #03#  CONNESSIONE = apertura e chiusura
'        #03.01# = Apri              #03.01.01# chiusura precedente  #03.01.02# corrente  #03.01.03# esterno
'        #03.02# = GetDatabase       (compatibile: apre il corrente se serve)
'        #03.03# = Chiudi
'  #04#  LETTURA = recordset, conteggi, singolo valore
'        #04.01# = ApriRecordset     #04.02# = ContaRecord    #04.03# = ValoreSQL
'  #05#  SCRITTURA = SQL di comando e query salvate
'        #05.01# = EseguiSQL         #05.01.01# database  #05.01.02# Execute
'        #05.02# = EseguiQuery       #05.02.01# esistenza  #05.02.02# Execute o DoCmd  #05.02.03# esito
'  #06#  CONTROLLI = esistenza di query, tabelle e campi
'        #06.01# = QueryEsiste       #06.02# = TabellaEsiste  #06.03# = CampoEsiste
'  #07#  ELENCO_QUERY = esecuzione di un elenco di query (ex ClsManager.AggiornaTuttiIDati)
'        #07.01# = EseguiElencoQuery #07.01.01# controlli  #07.01.02# ciclo  #07.01.03# report
'        #07.02# = AggiungiQuery     (aiuto per costruire l'elenco)
'  #08#  TRANSAZIONI = inizio, conferma, annullamento
'        #08.01# = IniziaTransazione #08.02# = ConfermaTransazione  #08.03# = AnnullaTransazione
'  #09#  COLLEGAMENTI = tabelle collegate
'        #09.01# = CollegaTabella    #09.01.01# esistente  #09.01.02# creazione
'        #09.02# = RicollegaTabelle  #09.02.01# filtro  #09.02.02# RefreshLink
'        #09.03# = ElencoCollegamenti
'  #10#  MANUTENZIONE = compattazione
'        #10.01# = CompattaDatabase
'  #90#  LOG_INTERNO = LOG_Scrivi, LOG_Errore, LOG_Passo, LOG_Cartella
'==================================================================================================


'##################################################################################################
'## #01# IMPOSTAZIONI = costanti e variabili
'##################################################################################################

'//#01.01# = costanti
Private Const NOME_CLASSE As String = "MSys_TB12_Cls03_DATABASE"

'//#01.02# = variabili della classe
Private m_Db As DAO.database
Private m_Esterno_b As Boolean
Private m_Percorso_s As String
Private m_RecordInteressati_l As Long
Private m_UltimoReport_s As String
Private m_InTransazione_b As Boolean
Private m_CartellaLog_s As String
Private m_LogAttivo_b As Boolean
Private m_LogDettagliato_b As Boolean
Private m_UltimoErrore_s As String


'##################################################################################################
'## #02# CICLO_DI_VITA_E_PROPRIETA = inizializzazione, chiusura, proprieta
'##################################################################################################

Private Sub Class_Initialize()
    '//#02.01# = nessun database aperto finche' non serve
    m_LogAttivo_b = True
    m_LogDettagliato_b = False
End Sub

Private Sub Class_Terminate()
    '//#02.02# = transazione aperta: annullata; database esterno: chiuso
    On Error Resume Next
    If m_InTransazione_b Then AnnullaTransazione
    Chiudi
End Sub

'//#02.03# = proprieta
Public Property Get DatabaseDAO() As DAO.database
    Set DatabaseDAO = GetDatabase()
End Property
Public Property Get Percorso() As String
    If m_Db Is Nothing Then Percorso = CurrentProject.FullName Else Percorso = m_Db.Name
End Property
Public Property Get RecordInteressati() As Long
    RecordInteressati = m_RecordInteressati_l
End Property
Public Property Get UltimoReport() As String
    UltimoReport = m_UltimoReport_s
End Property
Public Property Get CartellaLog() As String
    CartellaLog = LOG_Cartella()
End Property
Public Property Let CartellaLog(ByVal par_Cartella_s As String)
    m_CartellaLog_s = par_Cartella_s
End Property
Public Property Get LogAttivo() As Boolean
    LogAttivo = m_LogAttivo_b
End Property
Public Property Let LogAttivo(ByVal par_Valore_b As Boolean)
    m_LogAttivo_b = par_Valore_b
End Property
Public Property Get LogDettagliato() As Boolean
    LogDettagliato = m_LogDettagliato_b
End Property
Public Property Let LogDettagliato(ByVal par_Valore_b As Boolean)
    m_LogDettagliato_b = par_Valore_b
End Property
Public Property Get UltimoErrore() As String
    UltimoErrore = m_UltimoErrore_s
End Property


'##################################################################################################
'## #03# CONNESSIONE = apertura e chiusura
'##################################################################################################

Public Function Apri(Optional ByVal percorsoDb As String = vbNullString, Optional ByVal solaLettura As Boolean = False) As Boolean
'//#03.01# = senza percorso usa il database corrente; con percorso apre un .mdb/.accdb esterno
Const PROC As String = "Apri"
Dim sPos_s As String
On Error GoTo ERRORE
    '//#03.01.01# = chiudo l'eventuale database aperto prima
    sPos_s = "03.01.01"
    Chiudi
    If Len(percorsoDb) = 0 Then
        '//#03.01.02# = database corrente
        sPos_s = "03.01.02"
        Set m_Db = CurrentDb
        m_Esterno_b = False
    Else
        '//#03.01.03# = database esterno
        sPos_s = "03.01.03"
        Set m_Db = DBEngine.OpenDatabase(percorsoDb, False, solaLettura)
        m_Esterno_b = True
    End If
    m_Percorso_s = m_Db.Name
    Apri = True
    LOG_Scrivi PROC, sPos_s, "OK", IIf(m_Esterno_b, "esterno", "corrente") & IIf(solaLettura, " (sola lettura)", "") & ": " & m_Percorso_s
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & percorsoDb & "]"
    Set m_Db = Nothing
    Apri = False
    Resume USCITA
End Function

Public Function GetDatabase() As DAO.database
'//#03.02# = database in uso (apre il corrente se nessuno e' aperto)
    If m_Db Is Nothing Then Apri
    Set GetDatabase = m_Db
End Function

Public Sub Chiudi()
'//#03.03# = chiude il database esterno (il corrente non si chiude, si rilascia)
    On Error Resume Next
    If Not m_Db Is Nothing Then
        If m_Esterno_b Then m_Db.Close
        LOG_Passo "Chiudi", "03.03", "rilasciato: " & m_Percorso_s
    End If
    Set m_Db = Nothing
    m_Esterno_b = False
End Sub


'##################################################################################################
'## #04# LETTURA = recordset, conteggi, singolo valore
'##################################################################################################

Public Function ApriRecordset(ByVal sql As String, Optional ByVal db As DAO.database = Nothing, Optional ByVal tipo As Long = dbOpenDynaset) As DAO.Recordset
'//#04.01# = recordset su SQL, tabella o query; Nothing in caso di errore
Const PROC As String = "ApriRecordset"
On Error GoTo ERRORE
    If db Is Nothing Then Set db = GetDatabase()
    Set ApriRecordset = db.OpenRecordset(sql, tipo)
    LOG_Passo PROC, "04.01", "aperto: " & sql
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, "04.01", Err.Number, Err.Description & " [" & sql & "]"
    Set ApriRecordset = Nothing
    Resume USCITA
End Function

Public Function ContaRecord(ByVal origine As String, Optional ByVal criterio As String = vbNullString) As Long
'//#04.02# = numero di record di una tabella o query, con criterio WHERE facoltativo; -1 in caso di errore
Const PROC As String = "ContaRecord"
Dim rs As DAO.Recordset, sSql_s As String
On Error GoTo ERRORE
    sSql_s = "SELECT COUNT(*) FROM [" & origine & "]"
    If Len(criterio) > 0 Then sSql_s = sSql_s & " WHERE " & criterio
    Set rs = GetDatabase().OpenRecordset(sSql_s, dbOpenSnapshot)
    ContaRecord = Nz(rs.Fields(0).value, 0)
    rs.Close
USCITA:
    Set rs = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, "04.02", Err.Number, Err.Description & " [" & sSql_s & "]"
    ContaRecord = -1
    Resume USCITA
End Function

Public Function ValoreSQL(ByVal sql As String, Optional ByVal valoreSeVuoto As Variant = Null) As Variant
'//#04.03# = primo campo del primo record (come DLookup, ma anche su database esterni)
Const PROC As String = "ValoreSQL"
Dim rs As DAO.Recordset
On Error GoTo ERRORE
    ValoreSQL = valoreSeVuoto
    Set rs = GetDatabase().OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then ValoreSQL = Nz(rs.Fields(0).value, valoreSeVuoto)
    rs.Close
USCITA:
    Set rs = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, "04.03", Err.Number, Err.Description & " [" & sql & "]"
    Resume USCITA
End Function


'##################################################################################################
'## #05# SCRITTURA = SQL di comando e query salvate
'##################################################################################################

Public Function EseguiSQL(ByVal sql As String, Optional ByVal db As DAO.database = Nothing) As Boolean
'//#05.01# = INSERT, UPDATE, DELETE, CREATE...; RecordInteressati contiene le righe toccate
Const PROC As String = "EseguiSQL"
Dim sPos_s As String
On Error GoTo ERRORE
    '//#05.01.01# = database
    sPos_s = "05.01.01"
    If db Is Nothing Then Set db = GetDatabase()
    '//#05.01.02# = Execute con dbFailOnError (in caso di errore nessuna modifica parziale)
    sPos_s = "05.01.02"
    m_RecordInteressati_l = 0
    db.Execute sql, dbFailOnError
    m_RecordInteressati_l = db.RecordsAffected
    EseguiSQL = True
    LOG_Scrivi PROC, sPos_s, "OK", m_RecordInteressati_l & " record: " & sql
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & sql & "]"
    EseguiSQL = False
    Resume USCITA
End Function

Public Function EseguiQuery(ByVal nomeQuery As String, Optional ByVal usaDoCmd As Boolean = False) As Boolean
'//#05.02# = esegue una query di comando salvata.
'           usaDoCmd = True: DoCmd.OpenQuery senza avvisi, necessario se la query legge valori dalle form
'           (Forms!...): funziona solo sul database corrente.
Const PROC As String = "EseguiQuery"
Dim sPos_s As String, qd As DAO.QueryDef
On Error GoTo ERRORE
    '//#05.02.01# = esistenza
    sPos_s = "05.02.01"
    If Not QueryEsiste(nomeQuery) Then
        LOG_Scrivi PROC, sPos_s, "AVVISO", "query inesistente: " & nomeQuery
        m_UltimoErrore_s = "#" & sPos_s & "# query inesistente: " & nomeQuery
        GoTo USCITA
    End If
    '//#05.02.02# = esecuzione
    sPos_s = "05.02.02"
    m_RecordInteressati_l = 0
    If usaDoCmd And Not m_Esterno_b Then
        DoCmd.SetWarnings False
        DoCmd.OpenQuery nomeQuery, acViewNormal, acEdit
        DoCmd.SetWarnings True
    Else
        Set qd = GetDatabase().QueryDefs(nomeQuery)
        qd.Execute dbFailOnError
        m_RecordInteressati_l = qd.RecordsAffected
    End If
    '//#05.02.03# = esito
    sPos_s = "05.02.03"
    EseguiQuery = True
    LOG_Scrivi PROC, sPos_s, "OK", nomeQuery & IIf(usaDoCmd, " (DoCmd)", " - " & m_RecordInteressati_l & " record")
USCITA:
    Set qd = Nothing
    Exit Function
ERRORE:
    DoCmd.SetWarnings True
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & nomeQuery & "]"
    EseguiQuery = False
    Resume USCITA
End Function


'##################################################################################################
'## #06# CONTROLLI = esistenza di query, tabelle e campi (senza leggere MSysObjects)
'##################################################################################################

Public Function QueryEsiste(ByVal nomeQuery As String) As Boolean
'//#06.01#
Dim qd As DAO.QueryDef
    On Error Resume Next
    Set qd = GetDatabase().QueryDefs(nomeQuery)
    QueryEsiste = (Err.Number = 0)
    Set qd = Nothing
End Function

Public Function TabellaEsiste(ByVal nomeTabella As String) As Boolean
'//#06.02# = tabella locale o collegata
Dim td As DAO.TableDef
    On Error Resume Next
    Set td = GetDatabase().TableDefs(nomeTabella)
    TabellaEsiste = (Err.Number = 0)
    Set td = Nothing
End Function

Public Function CampoEsiste(ByVal nomeTabella As String, ByVal nomeCampo As String) As Boolean
'//#06.03#
Dim f As DAO.Field
    On Error Resume Next
    Set f = GetDatabase().TableDefs(nomeTabella).Fields(nomeCampo)
    CampoEsiste = (Err.Number = 0)
    Set f = Nothing
End Function


'##################################################################################################
'## #07# ELENCO_QUERY = esecuzione di un elenco di query (ex ClsManager.AggiornaTuttiIDati)
'##################################################################################################

Public Function EseguiElencoQuery(ByVal elenco As Collection, Optional ByVal mostraReport As Boolean = False, Optional ByVal usaDoCmd As Boolean = False) As Boolean
'//#07.01# = elenco di Array(NomeQuery, Descrizione); True se tutte eseguite. Report in UltimoReport
Const PROC As String = "EseguiElencoQuery"
Dim sPos_s As String, v As Variant, nOk_l As Long, nNo_l As Long, sRep_s As String, sDescr_s As String
On Error GoTo ERRORE
    '//#07.01.01# = controlli
    sPos_s = "07.01.01"
    If elenco Is Nothing Then GoTo USCITA
    LOG_Scrivi PROC, sPos_s, "INIZIO", elenco.Count & " query da eseguire"
    '//#07.01.02# = ciclo: le query mancanti o in errore non fermano le altre
    sPos_s = "07.01.02"
    For Each v In elenco
        sDescr_s = v(0)
        If UBound(v) >= 1 Then sDescr_s = v(1)
        If EseguiQuery(CStr(v(0)), usaDoCmd) Then
            nOk_l = nOk_l + 1
            sRep_s = sRep_s & "[OK] " & sDescr_s & vbCrLf
        Else
            nNo_l = nNo_l + 1
            sRep_s = sRep_s & "[NO] " & sDescr_s & " (" & v(0) & ")" & vbCrLf
        End If
    Next v
    '//#07.01.03# = report
    sPos_s = "07.01.03"
    m_UltimoReport_s = sRep_s & vbCrLf & "Eseguite: " & nOk_l & " - Non trovate o in errore: " & nNo_l
    EseguiElencoQuery = (nNo_l = 0)
    LOG_Scrivi PROC, sPos_s, IIf(nNo_l = 0, "OK", "AVVISO"), "eseguite " & nOk_l & ", non eseguite " & nNo_l
    If mostraReport Then MsgBox "OPERAZIONI COMPLETATE:" & vbCrLf & vbCrLf & m_UltimoReport_s, IIf(nNo_l = 0, vbInformation, vbExclamation), "Esecuzione query"
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description
    EseguiElencoQuery = False
    Resume USCITA
End Function

Public Sub AggiungiQuery(ByRef elenco As Collection, ByVal nomeQuery As String, Optional ByVal descrizione As String = vbNullString)
'//#07.02# = aggiunge una voce all'elenco (crea la Collection se serve)
    If elenco Is Nothing Then Set elenco = New Collection
    If Len(descrizione) = 0 Then descrizione = nomeQuery
    elenco.Add Array(nomeQuery, descrizione)
End Sub


'##################################################################################################
'## #08# TRANSAZIONI = inizio, conferma, annullamento (Workspace predefinito di DBEngine)
'##################################################################################################

Public Function IniziaTransazione() As Boolean
'//#08.01#
    On Error GoTo ERRORE
    GetDatabase
    DBEngine.Workspaces(0).BeginTrans
    m_InTransazione_b = True
    IniziaTransazione = True
    LOG_Scrivi "IniziaTransazione", "08.01", "OK", "transazione aperta"
    Exit Function
ERRORE:
    LOG_Errore "IniziaTransazione", "08.01", Err.Number, Err.Description
End Function

Public Function ConfermaTransazione() As Boolean
'//#08.02#
    On Error GoTo ERRORE
    DBEngine.Workspaces(0).CommitTrans
    m_InTransazione_b = False
    ConfermaTransazione = True
    LOG_Scrivi "ConfermaTransazione", "08.02", "OK", "modifiche confermate"
    Exit Function
ERRORE:
    LOG_Errore "ConfermaTransazione", "08.02", Err.Number, Err.Description
End Function

Public Function AnnullaTransazione() As Boolean
'//#08.03#
    On Error GoTo ERRORE
    DBEngine.Workspaces(0).Rollback
    m_InTransazione_b = False
    AnnullaTransazione = True
    LOG_Scrivi "AnnullaTransazione", "08.03", "OK", "modifiche annullate"
    Exit Function
ERRORE:
    LOG_Errore "AnnullaTransazione", "08.03", Err.Number, Err.Description
End Function


'##################################################################################################
'## #09# COLLEGAMENTI = tabelle collegate
'##################################################################################################

Public Function CollegaTabella(ByVal percorsoDb As String, ByVal tabellaSorgente As String, Optional ByVal nomeLocale As String = vbNullString) As Boolean
'//#09.01# = crea (o ricrea) il collegamento a una tabella di un altro database
Const PROC As String = "CollegaTabella"
Dim sPos_s As String, td As DAO.TableDef, db As DAO.database
On Error GoTo ERRORE
    If Len(nomeLocale) = 0 Then nomeLocale = tabellaSorgente
    Set db = GetDatabase()
    '//#09.01.01# = esiste gia': se e' un collegamento lo sostituisco, se e' una tabella locale mi fermo
    sPos_s = "09.01.01"
    If TabellaEsiste(nomeLocale) Then
        If Len(db.TableDefs(nomeLocale).Connect) = 0 Then
            LOG_Scrivi PROC, sPos_s, "ERRORE", "esiste una tabella LOCALE con lo stesso nome: " & nomeLocale
            m_UltimoErrore_s = "#" & sPos_s & "# tabella locale con lo stesso nome"
            GoTo USCITA
        End If
        db.TableDefs.Delete nomeLocale
    End If
    '//#09.01.02# = creazione del collegamento
    sPos_s = "09.01.02"
    Set td = db.CreateTableDef(nomeLocale)
    td.Connect = ";DATABASE=" & percorsoDb
    td.SourceTableName = tabellaSorgente
    db.TableDefs.Append td
    db.TableDefs.Refresh
    CollegaTabella = True
    LOG_Scrivi PROC, sPos_s, "OK", nomeLocale & " -> " & percorsoDb & " [" & tabellaSorgente & "]"
USCITA:
    Set td = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & tabellaSorgente & "]"
    CollegaTabella = False
    Resume USCITA
End Function

Public Function RicollegaTabelle(ByVal nuovoPercorso As String, Optional ByVal filtroNome As String = "*", Optional ByVal vecchioPercorso As String = vbNullString) As Long
'//#09.02# = punta a nuovoPercorso i collegamenti che corrispondono al filtro (e al vecchio percorso, se indicato).
'           Restituisce quanti collegamenti sono stati aggiornati; -1 in caso di errore
Const PROC As String = "RicollegaTabelle"
Dim sPos_s As String, td As DAO.TableDef, nTot_l As Long, sNomeTd_s As String
On Error GoTo ERRORE
    For Each td In GetDatabase().TableDefs
        '//#09.02.01# = filtro: solo tabelle collegate ad Access, nome e vecchio percorso
        sPos_s = "09.02.01"
        sNomeTd_s = td.Name
        If Left$(td.Connect, 10) = ";DATABASE=" And (td.Name Like filtroNome) Then
            If Len(vecchioPercorso) = 0 Or InStr(1, td.Connect, vecchioPercorso, vbTextCompare) > 0 Then
                '//#09.02.02# = nuovo percorso e RefreshLink
                sPos_s = "09.02.02"
                td.Connect = ";DATABASE=" & nuovoPercorso
                td.RefreshLink
                nTot_l = nTot_l + 1
                LOG_Passo PROC, sPos_s, td.Name & " -> " & nuovoPercorso
            End If
        End If
    Next td
    RicollegaTabelle = nTot_l
    LOG_Scrivi PROC, sPos_s, "OK", nTot_l & " collegamenti aggiornati verso " & nuovoPercorso
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & sNomeTd_s & "]"
    RicollegaTabelle = -1
    Resume USCITA
End Function

Public Function ElencoCollegamenti(Optional ByVal filtroNome As String = "*") As Collection
'//#09.03# = Collection di "nome|percorso del database"
Dim td As DAO.TableDef, coll As New Collection
    On Error Resume Next
    For Each td In GetDatabase().TableDefs
        If Left$(td.Connect, 10) = ";DATABASE=" And (td.Name Like filtroNome) Then coll.Add td.Name & "|" & Mid$(td.Connect, 11)
    Next td
    Set ElencoCollegamenti = coll
End Function


'##################################################################################################
'## #10# MANUTENZIONE = compattazione
'##################################################################################################

Public Function CompattaDatabase(ByVal percorsoSorgente As String, ByVal percorsoDestinazione As String) As Boolean
'//#10.01# = compatta un database CHIUSO in un nuovo file (la destinazione non deve esistere)
Const PROC As String = "CompattaDatabase"
On Error GoTo ERRORE
    DBEngine.CompactDatabase percorsoSorgente, percorsoDestinazione
    CompattaDatabase = True
    LOG_Scrivi PROC, "10.01", "OK", percorsoSorgente & " -> " & percorsoDestinazione
    Exit Function
ERRORE:
    LOG_Errore PROC, "10.01", Err.Number, Err.Description & " [" & percorsoSorgente & "]"
End Function


'##################################################################################################
'## #90# LOG_INTERNO = <db>\Logs\MSys_TB12_Cls03_DATABASE_AAAA_MM_GG.txt
'##################################################################################################

Private Sub LOG_Scrivi(ByVal sFunz As String, ByVal sPos As String, ByVal sEsito As String, ByVal sTesto As String)
'//#90.01# = la stessa riga in Debug.Print e nel file
Dim sRiga_s As String, iF_i As Integer
    If Not m_LogAttivo_b Then Exit Sub
    On Error Resume Next
    sRiga_s = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Classe: " & NOME_CLASSE & " | Funzione: " & sFunz & _
              " | Pos: " & Left$(sPos & Space$(8), 8) & " | Esito: " & Left$(sEsito & Space$(6), 6) & " | " & sTesto
    Debug.Print sRiga_s
    iF_i = FreeFile
    Open LOG_Cartella() & "\" & NOME_CLASSE & "_" & Format$(Date, "yyyy_mm_dd") & ".txt" For Append As #iF_i
    Print #iF_i, sRiga_s
    Close #iF_i
End Sub

Private Sub LOG_Errore(ByVal sFunz As String, ByVal sPos As String, ByVal nErr As Long, ByVal sDescr As String)
'//#90.02#
    m_UltimoErrore_s = "#" & sPos & "# " & sFunz & " - errore " & nErr & ": " & sDescr
    LOG_Scrivi sFunz, sPos, "ERRORE", "errore " & nErr & ": " & sDescr & " -> codice fermo a #" & sPos & "#"
End Sub

Private Sub LOG_Passo(ByVal sFunz As String, ByVal sPos As String, ByVal sTesto As String)
'//#90.03# = solo con LogDettagliato = True
    If m_LogDettagliato_b Then LOG_Scrivi sFunz, sPos, "INFO", sTesto
End Sub

Private Function LOG_Cartella() As String
'//#90.04# = cartella impostata oppure <cartella del database>\Logs (creata se manca)
    On Error Resume Next
    If Len(m_CartellaLog_s) > 0 Then LOG_Cartella = m_CartellaLog_s Else LOG_Cartella = CurrentProject.path & "\Logs"
    If Len(Dir$(LOG_Cartella, vbDirectory)) = 0 Then MkDir LOG_Cartella
End Function
