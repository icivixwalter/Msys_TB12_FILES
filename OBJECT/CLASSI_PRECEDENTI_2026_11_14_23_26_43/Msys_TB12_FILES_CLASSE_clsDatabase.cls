VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Msys_TB12_FILES_CLASSE_clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================
' CLASSE: Msys_TB12_FILES_CLASSE_clsDatabase
' RESPONSABILITA': Gestione centralizzata accesso al database
' =====================================================
Option Compare Database
Option Explicit

' 00_VARIABILI = Oggetti database
Private m_db As DAO.database
Private m_rs As DAO.Recordset

' =====================================================
' Metodo: GetDatabase
' Descrizione: Restituisce l'oggetto database corrente
' =====================================================
Public Function GetDatabase() As DAO.database
    On Error GoTo ErrHandler
    
    If m_db Is Nothing Then
        Set m_db = CurrentDb
    End If
    
    Set GetDatabase = m_db
    
    Exit Function
    
ErrHandler:
    Set GetDatabase = Nothing
End Function

' =====================================================
' Metodo: ApriRecordset
' Descrizione: Apre un recordset con gestione errori
' =====================================================
Public Function ApriRecordset(sql As String, Optional db As DAO.database = Nothing) As DAO.Recordset
    On Error GoTo ErrHandler
    
    Dim dbLocal As DAO.database
    If db Is Nothing Then
        Set dbLocal = GetDatabase()
    Else
        Set dbLocal = db
    End If
    
    Set ApriRecordset = dbLocal.OpenRecordset(sql)
    
    Exit Function
    
ErrHandler:
    Set ApriRecordset = Nothing
End Function

' =====================================================
' Metodo: EseguiSQL
' Descrizione: Esegue una query SQL (UPDATE, DELETE, INSERT)
' =====================================================
Public Function EseguiSQL(sql As String, Optional db As DAO.database = Nothing) As Boolean
    On Error GoTo ErrHandler
    
    Dim dbLocal As DAO.database
    If db Is Nothing Then
        Set dbLocal = GetDatabase()
    Else
        Set dbLocal = db
    End If
    
    dbLocal.Execute sql, dbFailOnError
    EseguiSQL = True
    
    Exit Function
    
ErrHandler:
    EseguiSQL = False
End Function

' =====================================================
' Metodo: Chiudi
' Descrizione: Rilascia le risorse
' =====================================================
Public Sub Chiudi()
    On Error Resume Next
    
    If Not m_rs Is Nothing Then
        m_rs.Close
        Set m_rs = Nothing
    End If
    
    Set m_db = Nothing
End Sub
