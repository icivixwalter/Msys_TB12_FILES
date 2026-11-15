VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSE_ClsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================
' CLASSE: DataManager
' RESPONSABILITA': Gestire l'esecuzione di query di aggiornamento
' =====================================================
Option Compare Database
Option Explicit

' Dipendenze (opzionali - possono essere Nothing se non disponibili)
Private m_Logger As Object  ' clsLogger (opzionale)
Private m_Inspector As Object  ' DatabaseObjectInspector (opzionale)

' Costruttore
Private Sub Class_Initialize()
    ' Tentativo di inizializzare i componenti opzionali
    On Error Resume Next
    Set m_Logger = CreateObject("clsLogger")
    Set m_Inspector = CreateObject("DatabaseObjectInspector")
    On Error GoTo 0
End Sub

' Distruttore
Private Sub Class_Terminate()
    Set m_Inspector = Nothing
    Set m_Logger = Nothing
End Sub

' =====================================================
' METODO PRINCIPALE: Aggiorna tutti i dati
' =====================================================
Public Function AggiornaTuttiIDati() As Boolean
    On Error GoTo ErrHandler
    
    Dim report As String
    Dim successCount As Integer
    Dim failCount As Integer
    
    ' Definisce tutte le query da eseguire (configurazione)
    Dim queries As Collection
    Set queries = GetQueryConfigurations()
    
    ' Esegue ogni query
    Dim i As Integer
    Dim qryNome As String
    Dim qryDesc As String
    
    For i = 1 To queries.Count
        qryNome = GetQueryNome(queries, i)
        qryDesc = GetQueryDescrizione(queries, i)
        
        If EseguiQuerySeEsiste(qryNome, qryDesc) Then
            successCount = successCount + 1
            report = report & "[OK] " & qryDesc & vbCrLf
        Else
            failCount = failCount + 1
            report = report & "[NO] " & qryDesc & " (query non trovata)" & vbCrLf
        End If
    Next i
    
    ' Mostra report finale
    MsgBox "OPERAZIONI COMPLETATE:" & vbCrLf & vbCrLf & report & vbCrLf & _
           "Eseguite: " & successCount & " - Non trovate: " & failCount, _
           vbInformation, "Aggiornamento Dati"
    
    AggiornaTuttiIDati = (failCount = 0)
    Exit Function
    
ErrHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "AggiornaTuttiIDati", Err.Number, Err.Description, Erl
    End If
    AggiornaTuttiIDati = False
End Function

' =====================================================
' Metodo: Esegue una query se esiste
' =====================================================
Private Function EseguiQuerySeEsiste(nomeQuery As String, Optional descrizione As String = "") As Boolean
    On Error GoTo ErrHandler
    
    ' Controllo se la query esiste usando MSysObjects (senza dipendenze esterne)
    If QueryEsiste(nomeQuery) Then
        DoCmd.OpenQuery nomeQuery, acViewNormal, acEdit
        EseguiQuerySeEsiste = True
        
        If descrizione <> "" Then
            Debug.Print "OK: Eseguita - " & descrizione
        End If
    Else
        If descrizione <> "" Then
            Debug.Print "NO: Non trovata - " & descrizione
        End If
        EseguiQuerySeEsiste = False
    End If
    
    Exit Function
    
ErrHandler:
    If Not m_Logger Is Nothing Then
        m_Logger.ScriviLogErrore "EseguiQuerySeEsiste", Err.Number, Err.Description, Erl
    End If
    EseguiQuerySeEsiste = False
End Function

' =====================================================
' Funzione: Verifica se una query esiste (usa MSysObjects)
' =====================================================
Private Function QueryEsiste(nomeQuery As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    sql = "SELECT Name FROM MSysObjects WHERE Type = 5 AND Name = '" & nomeQuery & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    QueryEsiste = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
    
ErrHandler:
    QueryEsiste = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' =====================================================
' CONFIGURAZIONE: Definisce tutte le query da eseguire
' =====================================================
Private Function GetQueryConfigurations() As Collection
    Dim coll As New Collection
    
    ' Formato: Aggiungi NomeQuery, Descrizione
    
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

' Helper: Estrae il nome della query da una configurazione
Private Function GetQueryNome(coll As Collection, index As Integer) As String
    Dim arr As Variant
    arr = coll(index)
    GetQueryNome = arr(0)
End Function

' Helper: Estrae la descrizione della query da una configurazione
Private Function GetQueryDescrizione(coll As Collection, index As Integer) As String
    Dim arr As Variant
    arr = coll(index)
    GetQueryDescrizione = arr(1)
End Function

