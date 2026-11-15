VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Msys_TB12_FILES_CLASSE_ClsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ====================================================================================================================
' CLASSE: Msys_TB12_FILES_CLASSE_ClsManager
' RESPONSABILITA': Gestire l'esecuzione di query di aggiornamento
' VERSIONE: 2.0 - Integrata con classi centralizzate
' ====================================================================================================================

Option Compare Database
Option Explicit


' ********************************************************************************************************************
' *** DIPENDENZE - CLASSI CENTRALIZZATE ***
' ********************************************************************************************************************

' 00_DIPENDENZA = Logger centralizzato
Private m_Logger As Msys_TB12_FILES_CLASSE_clsLogger

' 00_DIPENDENZA = Database centralizzato
Private m_Database As Msys_TB12_FILES_CLASSE_clsDatabase

' 00_DIPENDENZA = FileSystem centralizzato
Private m_FileSystem As Msys_TB12_FILES_CLASSE_clsFileSystem


' ********************************************************************************************************************
' *** METODI DI INIZIALIZZAZIONE ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: Class_Initialize
' Descrizione: Inizializza la classe e le dipendenze
' ====================================================================================================================
Private Sub Class_Initialize()
    On Error Resume Next
    
    ' 01_INIZIALIZZA = Logger (se disponibile)
    Set m_Logger = New Msys_TB12_FILES_CLASSE_clsLogger
    
    ' 02_INIZIALIZZA = Database
    Set m_Database = New Msys_TB12_FILES_CLASSE_clsDatabase
    
    ' 03_INIZIALIZZA = FileSystem
    Set m_FileSystem = New Msys_TB12_FILES_CLASSE_clsFileSystem
    
    On Error GoTo 0
End Sub

' ====================================================================================================================
' Metodo: Class_Terminate
' Descrizione: Rilascia le risorse
' ====================================================================================================================
Private Sub Class_Terminate()
    On Error Resume Next
    
    ' 01_PULIZIA = Logger
    Set m_Logger = Nothing
    
    ' 02_PULIZIA = Database
    If Not m_Database Is Nothing Then
        m_Database.Chiudi
        Set m_Database = Nothing
    End If
    
    ' 03_PULIZIA = FileSystem
    Set m_FileSystem = Nothing
End Sub

' ====================================================================================================================
' Metodo: Init
' Descrizione: Inizializza la classe con dipendenze esterne (opzionale)
' Parametri:
'   logger    - Istanza del logger (opzionale)
'   database  - Istanza del database (opzionale)
'   fileSystem - Istanza del fileSystem (opzionale)
' ====================================================================================================================
Public Sub Init(Optional logger As Msys_TB12_FILES_CLASSE_clsLogger = Nothing, _
                Optional database As Msys_TB12_FILES_CLASSE_clsDatabase = Nothing, _
                Optional fileSystem As Msys_TB12_FILES_CLASSE_clsFileSystem = Nothing)
    
    On Error Resume Next
    
    ' 01_INIT = Logger
    If Not logger Is Nothing Then
        Set m_Logger = logger
    ElseIf m_Logger Is Nothing Then
        Set m_Logger = New Msys_TB12_FILES_CLASSE_clsLogger
    End If
    
    ' 02_INIT = Database
    If Not database Is Nothing Then
        Set m_Database = database
    ElseIf m_Database Is Nothing Then
        Set m_Database = New Msys_TB12_FILES_CLASSE_clsDatabase
    End If
    
    ' 03_INIT = FileSystem
    If Not fileSystem Is Nothing Then
        Set m_FileSystem = fileSystem
    ElseIf m_FileSystem Is Nothing Then
        Set m_FileSystem = New Msys_TB12_FILES_CLASSE_clsFileSystem
    End If
    
    On Error GoTo 0
End Sub


' ********************************************************************************************************************
' *** METODO PRINCIPALE ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: AggiornaTuttiIDati
' Descrizione: Esegue tutte le query di aggiornamento configurate
' Output: True se tutte le query sono state eseguite con successo
' ====================================================================================================================
Public Function AggiornaTuttiIDati() As Boolean
    On Error GoTo ErrHandler
    
    ' 01_INIZIALIZZA = Variabili
    Dim report As String
    Dim successCount As Integer
    Dim failCount As Integer
    Dim skipCount As Integer
    
    ' 02_CONFIG = Query da eseguire
    Dim queries As Collection
    Set queries = GetQueryConfigurations()
    
    ' 03_LOG = Inizio operazioni
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogEvento "AggiornaTuttiIDati", "Avvio aggiornamento dati - " & queries.Count & " query da eseguire"
    End If
    
    ' 04_ESEGUI = Ogni query
    Dim i As Integer
    Dim qryNome As String
    Dim qryDesc As String
    Dim esito As Boolean
    
    For i = 1 To queries.Count
        qryNome = GetQueryNome(queries, i)
        qryDesc = GetQueryDescrizione(queries, i)
        
        esito = EseguiQuerySeEsiste(qryNome, qryDesc)
        
        If esito Then
            successCount = successCount + 1
            report = report & "[OK] " & qryDesc & vbCrLf
        Else
            failCount = failCount + 1
            report = report & "[NO] " & qryDesc & " (query non trovata o errore)" & vbCrLf
        End If
    Next i
    
    ' 05_LOG = Fine operazioni
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogEvento "AggiornaTuttiIDati", "Aggiornamento completato. OK: " & successCount & ", NO: " & failCount
    End If
    
    ' 06_REPORT = Mostra report finale
    MsgBox "OPERAZIONI COMPLETATE:" & vbCrLf & vbCrLf & report & vbCrLf & _
           "Eseguite: " & successCount & " - Non trovate/Errore: " & failCount, _
           vbInformation, "Aggiornamento Dati"
    
    ' 07_RISULTATO = True se tutto OK
    AggiornaTuttiIDati = (failCount = 0)
    Exit Function
    
ErrHandler:
    ' 08_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "AggiornaTuttiIDati", Err.Number, Err.Description, Erl
    End If
    AggiornaTuttiIDati = False
End Function


' ********************************************************************************************************************
' *** METODI PRIVATI PER L'ESECUZIONE DELLE QUERY ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: EseguiQuerySeEsiste
' Descrizione: Esegue una query se esiste nel database
' Parametri:
'   nomeQuery     - Nome della query da eseguire
'   descrizione   - Descrizione della query (per log)
' Output: True se la query è stata eseguita con successo
' ====================================================================================================================
Private Function EseguiQuerySeEsiste(nomeQuery As String, Optional descrizione As String = "") As Boolean
    On Error GoTo ErrHandler
    
    ' 01_VERIFICA = Query esiste?
    If QueryEsiste(nomeQuery) Then
        ' 02_ESEGUI = Apri la query
        DoCmd.OpenQuery nomeQuery, acViewNormal, acEdit
        EseguiQuerySeEsiste = True
        
        ' 03_LOG = Successo
        If descrizione <> "" Then
            Debug.Print "OK: Eseguita - " & descrizione
            If Not m_Logger Is Nothing Then
                m_Logger.ScriviLogEvento "EseguiQuerySeEsiste", "Query eseguita: " & descrizione & " (" & nomeQuery & ")"
            End If
        End If
    Else
        ' 04_LOG = Non trovata
        If descrizione <> "" Then
            Debug.Print "NO: Non trovata - " & descrizione
            If Not m_Logger Is Nothing Then
                m_Logger.ScriviLogEvento "EseguiQuerySeEsiste", "Query non trovata: " & descrizione & " (" & nomeQuery & ")"
            End If
        End If
        EseguiQuerySeEsiste = False
    End If
    
    Exit Function
    
ErrHandler:
    ' 05_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "EseguiQuerySeEsiste", Err.Number, Err.Description, Erl
    End If
    EseguiQuerySeEsiste = False
End Function

' ====================================================================================================================
' Funzione: QueryEsiste
' Descrizione: Verifica se una query esiste nel database
' Parametri:
'   nomeQuery     - Nome della query da verificare
' Output: True se la query esiste
' ====================================================================================================================
Private Function QueryEsiste(nomeQuery As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 01_DB = Usa il database centralizzato
    Dim rs As Object  ' DAO.Recordset con binding tardivo
    Dim sql As String
    
    ' 02_SQL = Query su MSysObjects
    sql = "SELECT Name FROM MSysObjects WHERE Type = 5 AND Name = '" & nomeQuery & "'"
    
    ' 03_APRI = Recordset
    If Not m_Database Is Nothing Then
        Set rs = m_Database.ApriRecordset(sql)
    Else
        ' Fallback se database non disponibile
        Dim db As Object
        Set db = CurrentDb()
        Set rs = db.OpenRecordset(sql)
    End If
    
    ' 04_VERIFICA = Esiste?
    QueryEsiste = Not rs.EOF
    
    ' 05_PULIZIA = Rilascia risorse
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    Exit Function
    
ErrHandler:
    QueryEsiste = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function


' ********************************************************************************************************************
' *** CONFIGURAZIONE DELLE QUERY ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: GetQueryConfigurations
' Descrizione: Definisce tutte le query da eseguire
' Output: Collection di array [NomeQuery, Descrizione]
' ====================================================================================================================
Private Function GetQueryConfigurations() As Collection
    Dim coll As New Collection
    
    ' --- PRES3000 (Calendario e Presenze) ---
    coll.Add Array("PRES3000_Qry01_10_AGGIORNA_Campo_IDKEY_CALENDARIO", "Aggiorna ID Key Calendario")
    coll.Add Array("PRES3000_Qry01_11_AGGIORNA_Campo_IDKEY_GIORNALIERE", "Aggiorna ID Key Giornaliere")
    coll.Add Array("PRES3000_Qry05_N01_ACCODA_Nw_ENTRATE->Orari", "Accoda Nuove Entrate Orari")
    coll.Add Array("PRES3000_Qry05_N02_ACCODA_Nw_USCITE->Orari", "Accoda Nuove Uscite Orari")
    
    ' --- INTESTATARI ---
    coll.Add Array("AGGIORNA_INTESTATARI_QyN01-01_Campo_Argomento", "Aggiorna Campo Argomento Intestatari")
    
    ' --- TABELLA ATTI ---
    coll.Add Array("LLPP_ATTI_Qy01-01_Gestione_AGGIORNA_IDKeyArch+IdOrd", "Aggiorna ID Key e ID Ordine Atti")
    coll.Add Array("LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFascicolo", "Aggiorna Denominazione Fascicolo")
    coll.Add Array("LLPP_ATTI_Qy01-11_AGGIORNA_ANNO", "Aggiorna Anno Atto (se nullo)")
    coll.Add Array("LLPP_ATTI_Qy01_01_Gestione_AGGIORNA_Campo_NroPratica_s=Null", "Aggiorna Numero Pratica (se nullo)")
    
    ' --- FALDONI E FASCICOLI ---
    coll.Add Array("LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFALDONE", "Aggiorna Denominazione Faldone")
    
    ' --- ALLEGATI ---
    coll.Add Array("LLPP_ATTI_Qy02_01_Allegati_AGGIORNA_CodOpera", "Aggiorna Codice Opera Allegati")
    coll.Add Array("LLPP_ATTI_Qry02_01_Allegati_AGGIORNA_IDGestione", "Aggiorna ID Gestione Allegati")
    
    ' --- AFFIDAMENTI ---
    coll.Add Array("AGGIORNA_AFFIDAMENTI_QyN10-01_Campi_TipoAtto+NroAtto+DataAtto", "Aggiorna Dati Affidamenti")
    
    ' --- ACCODAMENTI ---
    coll.Add Array("Accoda_QyDF15-01_NwEVENTI_Ins", "Accoda Nuovi Eventi")
    coll.Add Array("LLPP_QryDF01_CodiceOpera_ACCODA_CODICI_Ins", "Accoda Nuovi Codici Opera")
    
    Set GetQueryConfigurations = coll
End Function

' ====================================================================================================================
' Helper: Estrae il nome della query da una configurazione
' ====================================================================================================================
Private Function GetQueryNome(coll As Collection, index As Integer) As String
    Dim arr As Variant
    arr = coll(index)
    GetQueryNome = arr(0)
End Function

' ====================================================================================================================
' Helper: Estrae la descrizione della query da una configurazione
' ====================================================================================================================
Private Function GetQueryDescrizione(coll As Collection, index As Integer) As String
    Dim arr As Variant
    arr = coll(index)
    GetQueryDescrizione = arr(1)
End Function

