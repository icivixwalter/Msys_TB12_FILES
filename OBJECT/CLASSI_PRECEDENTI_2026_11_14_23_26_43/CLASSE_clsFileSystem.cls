VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSE_clsFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================
' CLASSE: CLASSE_clsFileSystem
' RESPONSABILITA': Gestione di file e cartelle (apertura, selezione, proprietà)
' SUPPORTO: Percorsi lunghi fino a 32.767 caratteri
' =====================================================
Option Compare Database
Option Explicit

' --- Costanti per ShellExecute ---
Private Const SW_SHOWNORMAL As Long = 1
Private Const SE_ERR_SUCCESS As Long = 32

' --- Costanti per GetOpenFileName ---
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_LONGNAMES As Long = &H200000

' --- Costanti per percorsi lunghi ---
Private Const MAX_PATH As Long = 260
Private Const LONG_PATH_PREFIX As String = "\\?\"

' --- Struttura per GetOpenFileName API ---
#If VBA7 Then
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As LongPtr
        lpTemplateName As String
        pvReserved As LongPtr
        dwReserved As Long
        FlagsEx As Long
    End Type
#Else
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If

' --- API Dichiarations ---
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
    
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
        pOpenfilename As OPENFILENAME) As Boolean
    
    Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" ( _
        pOpenfilename As OPENFILENAME) As Boolean
    
    Private Declare PtrSafe Function GetFileAttributesEx Lib "kernel32" Alias "GetFileAttributesExA" ( _
        ByVal lpFileName As String, _
        ByVal fInfoLevelId As Long, _
        ByRef lpFileInformation As Any) As Long
    
    Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" ( _
        ByVal lpPathName As String) As Long
    
    Private Declare PtrSafe Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
    
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
        pOpenfilename As OPENFILENAME) As Boolean
    
    Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" ( _
        pOpenfilename As OPENFILENAME) As Boolean
    
    Private Declare Function GetFileAttributesEx Lib "kernel32" Alias "GetFileAttributesExA" ( _
        ByVal lpFileName As String, _
        ByVal fInfoLevelId As Long, _
        ByRef lpFileInformation As Any) As Long
    
    Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" ( _
        ByVal lpPathName As String) As Long
    
    Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
#End If

' =====================================================
' METODI PUBBLICI
' =====================================================

' --- Seleziona un file (con supporto percorsi lunghi) ---
Public Function SelezionaFile(Optional filtroDescrizione As String = "Tutti i file (*.*)", _
                               Optional filtroEstensione As String = "*.*", _
                               Optional cartellaIniziale As String = vbNullString, _
                               Optional titolo As String = "Seleziona un file") As String
    
    On Error GoTo ErrHandler
    
    Dim ofn As OPENFILENAME
    Dim strFileName As String
    Dim strFilter As String
    
    ' Prepara il filtro
    strFilter = filtroDescrizione & Chr(0) & filtroEstensione & Chr(0) & Chr(0)
    
    ' Buffer per nome file (supporta percorsi lunghi)
    strFileName = Space(32767) ' Supporta fino a 32.767 caratteri
    
    ' Inizializza la struttura
    ofn.lStructSize = Len(ofn)
    
    #If VBA7 Then
        ofn.hwndOwner = Application.hWndAccessApp
    #Else
        ofn.hwndOwner = Application.hWndAccessApp
    #End If
    
    ofn.lpstrFilter = strFilter
    ofn.nFilterIndex = 1
    ofn.lpstrFile = strFileName
    ofn.nMaxFile = Len(strFileName)
    ofn.lpstrFileTitle = Space(32767)
    ofn.nMaxFileTitle = 32767
    ofn.lpstrInitialDir = cartellaIniziale
    ofn.lpstrTitle = titolo
    ofn.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_LONGNAMES Or OFN_EXPLORER
    
    ' Mostra la finestra di dialogo
    If GetOpenFileName(ofn) Then
        SelezionaFile = TrimNull(ofn.lpstrFile)
    Else
        SelezionaFile = vbNullString
    End If
    
    Exit Function
    
ErrHandler:
    SelezionaFile = vbNullString
End Function

' --- Salva un file (con supporto percorsi lunghi) ---
Public Function SalvaFile(Optional filtroDescrizione As String = "Tutti i file (*.*)", _
                           Optional filtroEstensione As String = "*.*", _
                           Optional cartellaIniziale As String = vbNullString, _
                           Optional nomeFileSuggerito As String = "", _
                           Optional titolo As String = "Salva file") As String
    
    On Error GoTo ErrHandler
    
    Dim ofn As OPENFILENAME
    Dim strFileName As String
    Dim strFilter As String
    
    ' Prepara il filtro
    strFilter = filtroDescrizione & Chr(0) & filtroEstensione & Chr(0) & Chr(0)
    
    ' Buffer per nome file (supporta percorsi lunghi)
    If nomeFileSuggerito <> "" Then
        strFileName = nomeFileSuggerito & Space(32767 - Len(nomeFileSuggerito))
    Else
        strFileName = Space(32767)
    End If
    
    ' Inizializza la struttura
    ofn.lStructSize = Len(ofn)
    
    #If VBA7 Then
        ofn.hwndOwner = Application.hWndAccessApp
    #Else
        ofn.hwndOwner = Application.hWndAccessApp
    #End If
    
    ofn.lpstrFilter = strFilter
    ofn.nFilterIndex = 1
    ofn.lpstrFile = strFileName
    ofn.nMaxFile = Len(strFileName)
    ofn.lpstrFileTitle = Space(32767)
    ofn.nMaxFileTitle = 32767
    ofn.lpstrInitialDir = cartellaIniziale
    ofn.lpstrTitle = titolo
    ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_LONGNAMES Or OFN_EXPLORER
    
    ' Mostra la finestra di dialogo
    If GetSaveFileName(ofn) Then
        SalvaFile = TrimNull(ofn.lpstrFile)
    Else
        SalvaFile = vbNullString
    End If
    
    Exit Function
    
ErrHandler:
    SalvaFile = vbNullString
End Function

' --- Apri un file o cartella (supporta percorsi lunghi) ---
Public Function ApriFile(percorsoCompleto As String, Optional showCmd As Long = SW_SHOWNORMAL) As Boolean
    On Error GoTo ErrHandler
    
    If Trim(percorsoCompleto) = "" Then
        ApriFile = False
        Exit Function
    End If
    
    ' Normalizza il percorso per supportare percorsi lunghi
    Dim percorsoShell As String
    percorsoShell = NormalizzaPercorso(percorsoCompleto)
    
    #If VBA7 Then
        Dim risultato As LongPtr
    #Else
        Dim risultato As Long
    #End If
    
    ' Per cartelle, usiamo "explore" invece di "open" per evitare problemi
    Dim operazione As String
    If CartellaEsiste(percorsoCompleto) Then
        operazione = "explore"
    Else
        operazione = "open"
    End If
    
    risultato = ShellExecute(0, operazione, percorsoShell, vbNullString, vbNullString, showCmd)
    ApriFile = (risultato > SE_ERR_SUCCESS)
    
    Exit Function
    
ErrHandler:
    ApriFile = False
End Function

' --- Verifica se un file esiste (supporta percorsi lunghi) ---
Public Function FileEsiste(filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    If Trim(filePath) = "" Then
        FileEsiste = False
        Exit Function
    End If
    
    Dim percorso As String
    percorso = NormalizzaPercorso(filePath)
    
    ' Usa Dir con percorsi lunghi
    If Len(percorso) > MAX_PATH Then
        ' Per percorsi lunghi, usa GetFileAttributesEx
        FileEsiste = FileEsisteLongPath(percorso)
    Else
        FileEsiste = (Dir(percorso) <> "")
    End If
    
    Exit Function
    
ErrHandler:
    FileEsiste = False
End Function

' --- Verifica se una cartella esiste (supporta percorsi lunghi) ---
Public Function CartellaEsiste(cartellaPath As String) As Boolean
    On Error GoTo ErrHandler
    
    If Trim(cartellaPath) = "" Then
        CartellaEsiste = False
        Exit Function
    End If
    
    Dim percorso As String
    percorso = NormalizzaPercorso(cartellaPath)
    
    ' Assicura che il percorso termini con backslash per la verifica
    If Right(percorso, 1) <> "\" Then
        percorso = percorso & "\"
    End If
    
    If Len(percorso) > MAX_PATH Then
        ' Per percorsi lunghi, usa GetFileAttributesEx
        CartellaEsiste = CartellaEsisteLongPath(percorso)
    Else
        CartellaEsiste = (Dir(percorso, vbDirectory) <> "")
    End If
    
    Exit Function
    
ErrHandler:
    CartellaEsiste = False
End Function

' --- Mostra le proprietà di un file (supporta percorsi lunghi) ---
Public Sub MostraProprietaFile(filePath As String)
    On Error GoTo ErrHandler
    
    If Trim(filePath) = "" Then
        MsgBox "Percorso file vuoto.", vbExclamation, "Errore"
        Exit Sub
    End If
    
    If Not FileEsiste(filePath) Then
        MsgBox "File non trovato: " & filePath, vbExclamation, "Errore"
        Exit Sub
    End If
    
    Dim fso As Object
    Dim file As Object
    Dim percorso As String
    
    ' Normalizza il percorso per FSO
    percorso = NormalizzaPercorsoFSO(filePath)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(percorso)
    
    Dim infoMsg As String
    infoMsg = "Proprietà del file:" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "Percorso: " & file.path & vbCrLf
    infoMsg = infoMsg & "Nome: " & file.Name & vbCrLf
    infoMsg = infoMsg & "Tipo: " & file.Type & vbCrLf
    infoMsg = infoMsg & "Dimensione: " & Format(file.size, "#,##0") & " bytes" & vbCrLf
    infoMsg = infoMsg & "Dimensione: " & Format(file.size / 1024, "#,##0.0") & " KB" & vbCrLf
    infoMsg = infoMsg & "Dimensione: " & Format(file.size / 1048576, "#,##0.0") & " MB" & vbCrLf
    infoMsg = infoMsg & "Creazione: " & Format(file.DateCreated, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    infoMsg = infoMsg & "Ultimo accesso: " & Format(file.DateLastAccessed, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    infoMsg = infoMsg & "Ultima modifica: " & Format(file.DateLastModified, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    infoMsg = infoMsg & "Attributi: " & GetFileAttributesText(filePath) & vbCrLf
    infoMsg = infoMsg & "Percorso lungo: " & IIf(Len(file.path) > MAX_PATH, "Sì", "No")
    
    MsgBox infoMsg, vbInformation, "Informazioni sul file"
    
    Set file = Nothing
    Set fso = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Errore: " & Err.Description, vbCritical, "Errore"
End Sub

' --- Copia un file (supporta percorsi lunghi) ---
Public Function CopiaFile(sourcePath As String, destinationPath As String, Optional overwrite As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    
    If Not FileEsiste(sourcePath) Then
        CopiaFile = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim src As String
    Dim dst As String
    src = NormalizzaPercorsoFSO(sourcePath)
    dst = NormalizzaPercorsoFSO(destinationPath)
    
    fso.CopyFile src, dst, overwrite
    CopiaFile = True
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    CopiaFile = False
    Set fso = Nothing
End Function

' --- Sposta/Rinomina un file (supporta percorsi lunghi) ---
Public Function SpostaFile(sourcePath As String, destinationPath As String) As Boolean
    On Error GoTo ErrHandler
    
    If Not FileEsiste(sourcePath) Then
        SpostaFile = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim src As String
    Dim dst As String
    src = NormalizzaPercorsoFSO(sourcePath)
    dst = NormalizzaPercorsoFSO(destinationPath)
    
    fso.MoveFile src, dst
    SpostaFile = True
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    SpostaFile = False
    Set fso = Nothing
End Function

' --- Elimina un file (supporta percorsi lunghi) ---
Public Function EliminaFile(filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    If Not FileEsiste(filePath) Then
        EliminaFile = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(filePath)
    
    fso.DeleteFile percorso
    EliminaFile = True
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    EliminaFile = False
    Set fso = Nothing
End Function

' --- Crea una cartella (supporta percorsi lunghi) ---
Public Function CreaCartella(cartellaPath As String) As Boolean
    On Error GoTo ErrHandler
    
    If CartellaEsiste(cartellaPath) Then
        CreaCartella = True
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(cartellaPath)
    
    ' Crea ricorsivamente se necessario
    Dim cartellaCorrente As String
    Dim parti() As String
    Dim i As Long
    
    ' Se il percorso è molto lungo, crea direttamente
    If Len(percorso) > MAX_PATH Then
        fso.CreateFolder percorso
    Else
        ' Crea passo passo per evitare problemi
        fso.CreateFolder percorso
    End If
    
    CreaCartella = True
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    CreaCartella = False
    Set fso = Nothing
End Function

' --- Elimina una cartella (supporta percorsi lunghi) ---
Public Function EliminaCartella(cartellaPath As String, Optional forceDelete As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    If Not CartellaEsiste(cartellaPath) Then
        EliminaCartella = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(cartellaPath)
    
    fso.DeleteFolder percorso, forceDelete
    EliminaCartella = True
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    EliminaCartella = False
    Set fso = Nothing
End Function

' --- Restituisce il nome del file da un percorso completo ---
Public Function GetFileName(percorsoCompleto As String) As String
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(percorsoCompleto)
    
    If FileEsiste(percorsoCompleto) Then
        GetFileName = fso.GetFileName(percorso)
    Else
        GetFileName = ""
    End If
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    GetFileName = ""
    Set fso = Nothing
End Function

' --- Restituisce il percorso della cartella ---
Public Function GetFolderPath(percorsoCompleto As String) As String
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(percorsoCompleto)
    
    If FileEsiste(percorsoCompleto) Then
        GetFolderPath = fso.GetParentFolderName(percorso)
    Else
        GetFolderPath = ""
    End If
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    GetFolderPath = ""
    Set fso = Nothing
End Function

' --- Restituisce l'estensione del file ---
Public Function GetFileExtension(percorsoCompleto As String) As String
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(percorsoCompleto)
    
    If FileEsiste(percorsoCompleto) Then
        GetFileExtension = fso.GetExtensionName(percorso)
    Else
        GetFileExtension = ""
    End If
    
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    GetFileExtension = ""
    Set fso = Nothing
End Function

' --- Restituisce la dimensione di un file in formato leggibile ---
Public Function GetFileSizeFormatted(filePath As String) As String
    On Error GoTo ErrHandler
    
    If Not FileEsiste(filePath) Then
        GetFileSizeFormatted = ""
        Exit Function
    End If
    
    Dim fso As Object
    Dim file As Object
    Dim size As Double
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(filePath)
    Set file = fso.GetFile(percorso)
    size = file.size
    
    Select Case size
        Case Is < 1024
            GetFileSizeFormatted = Format(size, "0") & " bytes"
        Case Is < 1048576
            GetFileSizeFormatted = Format(size / 1024, "0.0") & " KB"
        Case Is < 1073741824
            GetFileSizeFormatted = Format(size / 1048576, "0.0") & " MB"
        Case Else
            GetFileSizeFormatted = Format(size / 1073741824, "0.0") & " GB"
    End Select
    
    Set file = Nothing
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    GetFileSizeFormatted = ""
    Set file = Nothing
    Set fso = Nothing
End Function

' --- Verifica se un percorso è lungo (> 260 caratteri) ---
Public Function IsLongPath(percorso As String) As Boolean
    IsLongPath = (Len(percorso) > MAX_PATH)
End Function

' --- Normalizza un percorso per l'uso con le API ---
Public Function NormalizzaPercorso(percorso As String) As String
    On Error GoTo ErrHandler
    
    If Trim(percorso) = "" Then
        NormalizzaPercorso = ""
        Exit Function
    End If
    
    Dim risultato As String
    risultato = percorso
    
    ' Rimuovi spazi iniziali e finali
    risultato = Trim(risultato)
    
    ' Se il percorso è lungo o contiene caratteri speciali, aggiungi il prefisso
    If Len(risultato) > MAX_PATH Then
        ' Se non ha già il prefisso \\?\
        If Left(risultato, 4) <> LONG_PATH_PREFIX Then
            ' Se è un percorso UNC (inizia con \\)
            If Left(risultato, 2) = "\\" Then
                ' Converti UNC in formato lungo: \\?\UNC\server\share\...
                risultato = "\\?\UNC\" & Mid(risultato, 3)
            Else
                ' Percorso normale con prefisso
                risultato = LONG_PATH_PREFIX & risultato
            End If
        End If
    End If
    
    NormalizzaPercorso = risultato
    
    Exit Function
    
ErrHandler:
    NormalizzaPercorso = percorso
End Function

' --- Normalizza un percorso per FSO ---
Private Function NormalizzaPercorsoFSO(percorso As String) As String
    On Error GoTo ErrHandler
    
    If Trim(percorso) = "" Then
        NormalizzaPercorsoFSO = ""
        Exit Function
    End If
    
    Dim risultato As String
    risultato = percorso
    
    ' FSO non supporta il prefisso \\?\ direttamente
    ' Rimuoviamo il prefisso se presente
    If Left(risultato, 4) = LONG_PATH_PREFIX Then
        risultato = Mid(risultato, 5)
        ' Se era un percorso UNC lungo, converti
        If Left(risultato, 4) = "UNC\" Then
            risultato = "\\" & Mid(risultato, 5)
        End If
    End If
    
    NormalizzaPercorsoFSO = risultato
    
    Exit Function
    
ErrHandler:
    NormalizzaPercorsoFSO = percorso
End Function

' =====================================================
' FUNZIONI PRIVATE
' =====================================================

' --- Rimuove i caratteri null dalla stringa ---
Private Function TrimNull(str As String) As String
    Dim pos As Integer
    pos = InStr(str, Chr(0))
    If pos > 0 Then
        TrimNull = Left(str, pos - 1)
    Else
        TrimNull = str
    End If
End Function

' --- Verifica esistenza file con percorso lungo (API) ---
Private Function FileEsisteLongPath(filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim attribs As Long
    attribs = GetFileAttributesExLong(filePath)
    FileEsisteLongPath = (attribs <> -1 And (attribs And 16) = 0)
    
    Exit Function
    
ErrHandler:
    FileEsisteLongPath = False
End Function

' --- Verifica esistenza cartella con percorso lungo (API) ---
Private Function CartellaEsisteLongPath(cartellaPath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim attribs As Long
    attribs = GetFileAttributesExLong(cartellaPath)
    CartellaEsisteLongPath = (attribs <> -1 And (attribs And 16) <> 0)
    
    Exit Function
    
ErrHandler:
    CartellaEsisteLongPath = False
End Function

' --- Ottiene attributi file via API (supporto percorsi lunghi) ---
Private Function GetFileAttributesExLong(filePath As String) As Long
    On Error GoTo ErrHandler
    
    ' Struttura per GetFileAttributesEx
    Dim dwFileAttributes As Long
    Dim ftCreationTime(1) As Currency
    Dim ftLastAccessTime(1) As Currency
    Dim ftLastWriteTime(1) As Currency
    Dim nFileSizeHigh As Long
    Dim nFileSizeLow As Long
    
    ' Usa la versione con prefisso lungo
    Dim percorso As String
    percorso = NormalizzaPercorso(filePath)
    
    ' Chiamata API
    If GetFileAttributesEx(percorso, 0, dwFileAttributes) <> 0 Then
        GetFileAttributesExLong = dwFileAttributes
    Else
        GetFileAttributesExLong = -1
    End If
    
    Exit Function
    
ErrHandler:
    GetFileAttributesExLong = -1
End Function

' --- Restituisce gli attributi del file in formato testo ---
Private Function GetFileAttributesText(filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Dim file As Object
    Dim attribs As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim percorso As String
    percorso = NormalizzaPercorsoFSO(filePath)
    Set file = fso.GetFile(percorso)
    
    attribs = ""
    If file.Attributes And 1 Then attribs = attribs & "Sola lettura, "
    If file.Attributes And 2 Then attribs = attribs & "Nascosto, "
    If file.Attributes And 4 Then attribs = attribs & "Sistema, "
    If file.Attributes And 16 Then attribs = attribs & "Directory, "
    If file.Attributes And 32 Then attribs = attribs & "Archivio, "
    If file.Attributes And 64 Then attribs = attribs & "Alias, "
    If file.Attributes And 128 Then attribs = attribs & "Compresso, "
    
    If attribs = "" Then
        GetFileAttributesText = "Normale"
    Else
        GetFileAttributesText = Left(attribs, Len(attribs) - 2)
    End If
    
    Set file = Nothing
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    GetFileAttributesText = "Non disponibile"
    Set file = Nothing
    Set fso = Nothing
End Function

