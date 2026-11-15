VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Msys_TB12_FILES_CLASSE_clsDBInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ====================================================================================================================
' CLASSE: Msys_TB12_FILES_CLASSE_clsDBInspector
' RESPONSABILITA': Gestione accodamento oggetti di sistema (tabelle, query, forms)
' ATTENZIONE: Questa classe manipola tabelle di sistema - Usare con cautela!
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


' ********************************************************************************************************************
' *** COSTANTI - NOMI QUERY ***
' ********************************************************************************************************************

' 00_COSTANTI = Query di accodamento
Private Const QRY_ACC_TABELLE As String = "MSys_TABELLE_Qry01_71_INSERT_toMsys_DF12"
Private Const QRY_ACC_QUERY As String = "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12"
Private Const QRY_ACC_FORMS As String = "MSys_FORMS_Qry01_71_INSERT_toMsys_DF12"


' ********************************************************************************************************************
' *** METODI DI INIZIALIZZAZIONE ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: Class_Initialize
' Descrizione: Inizializza la classe
' ====================================================================================================================
Private Sub Class_Initialize()
    On Error Resume Next
    
    ' 01_INIZIALIZZA = Logger (se disponibile)
    Set m_Logger = New Msys_TB12_FILES_CLASSE_clsLogger
    
    ' 02_INIZIALIZZA = Database
    Set m_Database = New Msys_TB12_FILES_CLASSE_clsDatabase
    
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
End Sub

' ====================================================================================================================
' Metodo: Init
' Descrizione: Inizializza la classe con dipendenze esterne
' Parametri:
'   logger    - Istanza del logger (opzionale)
'   database  - Istanza del database (opzionale)
' ====================================================================================================================
Public Sub Init(Optional logger As Msys_TB12_FILES_CLASSE_clsLogger = Nothing, _
                Optional database As Msys_TB12_FILES_CLASSE_clsDatabase = Nothing)
    
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
    
    On Error GoTo 0
End Sub


' ********************************************************************************************************************
' *** METODO PRINCIPALE ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: AccodaTuttiGliOggetti
' Descrizione: Accoda tutti gli oggetti (tabelle, query, forms)
' Parametri:
'   codiceProgetto  - Codice del progetto da associare agli oggetti
' Output: True se tutte le operazioni sono riuscite
' ====================================================================================================================
Public Function AccodaTuttiGliOggetti(ByVal codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 01_LOG = Inizio operazioni
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogEvento "AccodaTuttiGliOggetti", "Avvio accodamento oggetti per progetto: " & codiceProgetto
    End If
    
    ' 02_INIZIALIZZA = Variabili
    Dim success As Boolean
    success = True
    Dim report As String
    report = ""
    
    ' 03_ESEGUI = Accodamento per ogni tipo
    ' Tabelle
    If AccodaTabelle(codiceProgetto) Then
        report = report & "[OK] Tabelle" & vbCrLf
    Else
        report = report & "[NO] Tabelle" & vbCrLf
        success = False
    End If
    
    ' Query
    If AccodaQuery(codiceProgetto) Then
        report = report & "[OK] Query" & vbCrLf
    Else
        report = report & "[NO] Query" & vbCrLf
        success = False
    End If
    
    ' Forms
    If AccodaForms(codiceProgetto) Then
        report = report & "[OK] Forms" & vbCrLf
    Else
        report = report & "[NO] Forms" & vbCrLf
        success = False
    End If
    
    ' 04_LOG = Fine operazioni
    If Not m_Logger Is Nothing Then
        If success Then
            m_Logger.ScriviLogEvento "AccodaTuttiGliOggetti", "Accodamento completato con successo"
        Else
            m_Logger.ScriviLogEvento "AccodaTuttiGliOggetti", "Accodamento completato con errori"
        End If
    End If
    
    ' 05_REPORT = Mostra risultato
    MsgBox "ACCODAMENTO OGGETTI COMPLETATO:" & vbCrLf & vbCrLf & report & vbCrLf & _
           "Risultato: " & IIf(success, "TUTTO OK", "QUALCHE ERRORE"), _
           IIf(success, vbInformation, vbExclamation), "Accodamento Oggetti"
    
    ' 06_RISULTATO = True se tutto OK
    AccodaTuttiGliOggetti = success
    Exit Function
    
ErrHandler:
    ' 07_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "AccodaTuttiGliOggetti", Err.Number, Err.Description, Erl
    End If
    AccodaTuttiGliOggetti = False
End Function


' ********************************************************************************************************************
' *** METODI PRIVATI PER ACCODAMENTO ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: AccodaTabelle
' Descrizione: Accoda le tabelle
' Parametri:
'   codiceProgetto  - Codice del progetto
' Output: True se l'operazione è riuscita
' ====================================================================================================================
Private Function AccodaTabelle(codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 01_ESEGUI = Query di accodamento tabelle
    If EseguiQueryAccodamento(QRY_ACC_TABELLE, "Tabelle") Then
        AccodaTabelle = True
    Else
        AccodaTabelle = False
    End If
    
    Exit Function
    
ErrHandler:
    ' 02_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "AccodaTabelle", Err.Number, Err.Description, Erl
    End If
    AccodaTabelle = False
End Function

' ====================================================================================================================
' Metodo: AccodaQuery
' Descrizione: Accoda le query
' Parametri:
'   codiceProgetto  - Codice del progetto
' Output: True se l'operazione è riuscita
' ====================================================================================================================
Private Function AccodaQuery(codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 01_ESEGUI = Query di accodamento query
    If EseguiQueryAccodamento(QRY_ACC_QUERY, "Query") Then
        AccodaQuery = True
    Else
        AccodaQuery = False
    End If
    
    Exit Function
    
ErrHandler:
    ' 02_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "AccodaQuery", Err.Number, Err.Description, Erl
    End If
    AccodaQuery = False
End Function

' ====================================================================================================================
' Metodo: AccodaForms
' Descrizione: Accoda i forms
' Parametri:
'   codiceProgetto  - Codice del progetto
' Output: True se l'operazione è riuscita
' ====================================================================================================================
Private Function AccodaForms(codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 01_ESEGUI = Query di accodamento forms
    If EseguiQueryAccodamento(QRY_ACC_FORMS, "Forms") Then
        AccodaForms = True
    Else
        AccodaForms = False
    End If
    
    Exit Function
    
ErrHandler:
    ' 02_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "AccodaForms", Err.Number, Err.Description, Erl
    End If
    AccodaForms = False
End Function

' ====================================================================================================================
' Metodo: EseguiQueryAccodamento
' Descrizione: Esegue una query di accodamento con verifica esistenza
' Parametri:
'   nomeQuery     - Nome della query da eseguire
'   tipoOggetto   - Tipo di oggetto (per log)
' Output: True se l'operazione è riuscita
' ====================================================================================================================
Private Function EseguiQueryAccodamento(nomeQuery As String, tipoOggetto As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 01_VERIFICA = Query esiste?
    If Not QueryEsiste(nomeQuery) Then
        ' 02_LOG = Query non trovata
        If Not m_Logger Is Nothing Then
            m_Logger.ScriviLogEvento "EseguiQueryAccodamento", "Query non trovata: " & nomeQuery & " (" & tipoOggetto & ")"
        End If
        EseguiQueryAccodamento = False
        Exit Function
    End If
    
    ' 03_ESEGUI = Apri la query
    DoCmd.OpenQuery nomeQuery, acViewNormal, acEdit
    EseguiQueryAccodamento = True
    
    ' 04_LOG = Successo
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogEvento "EseguiQueryAccodamento", "Accodamento eseguito: " & tipoOggetto & " (" & nomeQuery & ")"
    End If
    
    Exit Function
    
ErrHandler:
    ' 05_LOG = Errore
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "EseguiQueryAccodamento", Err.Number, Err.Description, Erl
    End If
    EseguiQueryAccodamento = False
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
' *** METODI DI UTILITY ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: GetConfigurazioneQuery
' Descrizione: Restituisce la configurazione delle query di accodamento
' Output: Collection con le query configurate
' ====================================================================================================================
Public Function GetConfigurazioneQuery() As Collection
    Dim coll As New Collection
    
    ' 01_CONFIG = Aggiungi tutte le query
    coll.Add Array(QRY_ACC_TABELLE, "Tabelle")
    coll.Add Array(QRY_ACC_QUERY, "Query")
    coll.Add Array(QRY_ACC_FORMS, "Forms")
    
    Set GetConfigurazioneQuery = coll
End Function

' ====================================================================================================================
' Metodo: VerificaTutteQuery
' Descrizione: Verifica se tutte le query di accodamento esistono
' Output: Collection con i risultati [NomeQuery, Esiste]
' ====================================================================================================================
Public Function VerificaTutteQuery() As Collection
    Dim coll As New Collection
    Dim config As Collection
    Dim i As Integer
    Dim nomeQuery As String
    
    ' 01_CONFIG = Ottieni configurazione
    Set config = GetConfigurazioneQuery()
    
    ' 02_VERIFICA = Ogni query
    For i = 1 To config.Count
        nomeQuery = config(i)(0)
        coll.Add Array(nomeQuery, QueryEsiste(nomeQuery))
    Next i
    
    Set VerificaTutteQuery = coll
End Function

