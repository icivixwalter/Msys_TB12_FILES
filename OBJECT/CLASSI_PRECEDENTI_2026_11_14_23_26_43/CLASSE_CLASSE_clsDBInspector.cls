VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSE_CLASSE_clsDBInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================
' CLASSE: clsSystemTableManager (use with caution!)
' ATTENZIONE: Questa classe manipola tabelle di sistema
' =============================================
Option Compare Database
Option Explicit

' NOTA: Non usa clsLogger per evitare dipendenze circolari

Private Sub Class_Initialize()
    ' Nessuna inizializzazione necessaria
End Sub

Private Sub Class_Terminate()
    ' Nessuna pulizia necessaria
End Sub

' =====================================================
' METODO PRINCIPALE: Accoda tutti gli oggetti
' =====================================================
Public Function AccodaTuttiGliOggetti(ByVal codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim success As Boolean
    success = True
    
    ' Esegue l accodamento per ogni tipo di oggetto
    If Not AccodaTabelle(codiceProgetto) Then
        LogErrore "AccodaTuttiGliOggetti", 1001, "Fallito accodamento tabelle"
        success = False
    End If
    
    If Not AccodaQuery(codiceProgetto) Then
        LogErrore "AccodaTuttiGliOggetti", 1002, "Fallito accodamento query"
        success = False
    End If
    
    If Not AccodaForms(codiceProgetto) Then
        LogErrore "AccodaTuttiGliOggetti", 1003, "Fallito accodamento forms"
        success = False
    End If
    
    AccodaTuttiGliOggetti = success
    Exit Function
    
ErrHandler:
    LogErrore "AccodaTuttiGliOggetti", Err.Number, Err.Description
    AccodaTuttiGliOggetti = False
End Function

' =====================================================
' Accoda Tabelle
' =====================================================
Private Function AccodaTabelle(codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    ' SOSTITUISCI con i nomi reali delle tue query
    DoCmd.OpenQuery "MSys_TABELLE_Qry01_71_INSERT_toMsys_DF12"
    
    AccodaTabelle = True
    Exit Function
    
ErrHandler:
    LogErrore "AccodaTabelle", Err.Number, Err.Description
    AccodaTabelle = False
End Function

' =====================================================
' Accoda Query
' =====================================================
Private Function AccodaQuery(codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    DoCmd.OpenQuery "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12"
    
    AccodaQuery = True
    Exit Function
    
ErrHandler:
    LogErrore "AccodaQuery", Err.Number, Err.Description
    AccodaQuery = False
End Function

' =====================================================
' Accoda Forms
' =====================================================
Private Function AccodaForms(codiceProgetto As String) As Boolean
    On Error GoTo ErrHandler
    
    DoCmd.OpenQuery "MSys_FORMS_Qry01_71_INSERT_toMsys_DF12"
    
    AccodaForms = True
    Exit Function
    
ErrHandler:
    LogErrore "AccodaForms", Err.Number, Err.Description
    AccodaForms = False
End Function

' =====================================================
' LOG ERRORI (interno, senza dipendenze)
' =====================================================
Private Sub LogErrore(nomeFunzione As String, codiceErrore As Long, descrizione As String)
    On Error Resume Next
    
    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim cartellaLogs As String
    
    cartellaLogs = CurrentProject.path & "\Logs"
    
    If Dir(cartellaLogs, vbDirectory) = "" Then
        MkDir cartellaLogs
    End If
    
    percorsoLog = cartellaLogs & "\SystemTableManager_Log.txt"
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        Print #fileNum, Format(Now, "yyyy-mm-dd HH:nn:ss") & " | " & nomeFunzione & " | Errore " & codiceErrore & ": " & descrizione
    Close #fileNum
End Sub
