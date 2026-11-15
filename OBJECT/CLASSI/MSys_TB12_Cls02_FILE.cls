VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSys_TB12_Cls02_FILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'==================================================================================================
'  CLASSE    : MSys_TB12_Cls02_FILE
'  PROGETTO  : MSys_TB12_FILES  (mini-progetto TB12 - gestione file e cartelle)
'  SCOPO     : file e cartelle con API Windows UNICODE a 32 e 64 bit, percorsi fino a 32.767
'              caratteri (prefisso \\?\), attributi, date, dimensioni, elenchi, apertura con
'              l'applicazione predefinita, finestre di selezione (Application.FileDialog).
'  SOSTITUISCE: Msys_TB12_FILES_CLASSE_clsFileSystem e CLASSE_clsFileSystem (stessi nomi pubblici)
'  DIPENDENZE: nessuna. La classe si importa da sola in qualsiasi progetto Access.
'  LOG       : <cartella del database>\Logs\MSys_TB12_Cls02_FILE_AAAA_MM_GG.txt
'              riga = data ora | Classe | Funzione | Pos: SS.PP.ss | Esito | testo  (identica in Debug.Print)
'  VERSIONE  : 1.0 - 2026-11-15
'--------------------------------------------------------------------------------------------------
'  INDICE DELLE SEZIONI, DEI PARAGRAFI E DEI SOTTOPARAGRAFI  (cerca l'etichetta, es. #05.01.03#)
'--------------------------------------------------------------------------------------------------
'  #01#  IMPOSTAZIONI = costanti, strutture e dichiarazioni API 32/64 bit, variabili
'        #01.01# = costanti generali, attributi e codici di errore Windows
'        #01.02# = strutture API: FILETIME, SYSTEMTIME, WIN32_FILE_ATTRIBUTE_DATA, WIN32_FIND_DATA
'        #01.03# = dichiarazioni API Unicode per VBA7 (Office 2010 e successivi, 32 e 64 bit)
'        #01.04# = dichiarazioni API Unicode per VBA6 (Office 2007 e precedenti, 32 bit)
'        #01.05# = variabili della classe
'  #02#  CICLO_DI_VITA_E_PROPRIETA = inizializzazione, chiusura, impostazioni del log
'        #02.01# = Class_Initialize: valori predefiniti
'        #02.02# = Class_Terminate
'        #02.03# = proprieta: CartellaLog, LogAttivo, LogDettagliato, UltimoErrore
'  #03#  VERIFICHE = esistenza di file e cartelle, percorso lungo
'        #03.01# = FileEsiste           #03.01.01# percorso vuoto  #03.01.02# lettura attributi  #03.01.03# esito
'        #03.02# = CartellaEsiste       #03.02.01# percorso vuoto  #03.02.02# lettura attributi  #03.02.03# esito
'        #03.03# = IsLongPath           percorso oltre 260 caratteri
'  #04#  PERCORSI = normalizzazione, prefisso \\?\, percorso breve 8.3, parti del percorso
'        #04.01# = NormalizzaPercorso   #04.01.01# pulizia  #04.01.02# UNC  #04.01.03# unita' locale
'        #04.02# = PercorsoSenzaPrefisso
'        #04.03# = PercorsoBreve        #04.03.01# chiamata GetShortPathNameW  #04.03.02# risultato
'        #04.04# = GetFileName, GetFolderPath, GetFileExtension, UnisciPercorso (solo testo)
'  #05#  CARTELLE = creazione ricorsiva, eliminazione, elenchi
'        #05.01# = CreaCartella         #05.01.01# gia' esiste  #05.01.02# radice  #05.01.03# livelli  #05.01.04# esito
'        #05.02# = EliminaCartella      #05.02.01# controlli  #05.02.02# contenuto (forzata)  #05.02.03# rimozione
'        #05.03# = ElencoFile           #05.03.01# controlli  #05.03.02# scansione  #05.03.03# esito
'        #05.04# = ElencoCartelle       #05.04.01# controlli  #05.04.02# scansione  #05.04.03# esito
'  #06#  FILE = copia, spostamento, eliminazione
'        #06.01# = CopiaFile            #06.01.01# controlli  #06.01.02# destinazione  #06.01.03# CopyFileW
'        #06.02# = SpostaFile           #06.02.01# controlli  #06.02.02# destinazione  #06.02.03# MoveFileExW
'        #06.03# = EliminaFile          #06.03.01# controlli  #06.03.02# sola lettura  #06.03.03# DeleteFileW
'  #07#  ATTRIBUTI_E_PROPRIETA = attributi, dimensione, date, scheda proprieta'
'        #07.01# = AttributiFile, AttributiTesto
'        #07.02# = DimensioneFile, GetFileSizeFormatted
'        #07.03# = DataCreazione, DataModifica, DataUltimoAccesso
'        #07.04# = ProprietaFile        #07.04.01# lettura  #07.04.02# composizione testo
'        #07.05# = MostraProprietaFile  (finestra di messaggio)
'  #08#  APERTURA = apertura con l'applicazione predefinita (anche oltre 260 caratteri)
'        #08.01# = ApriFile             #08.01.01# controlli  #08.01.02# percorso per la shell  #08.01.03# ShellExecuteW
'        #08.02# = ApriCartella
'  #09#  FINESTRE_DI_DIALOGO = selezione file, file multipli, cartella, salvataggio
'        #09.01# = SelezionaFile        #09.02# = SelezionaFileMultipli
'        #09.03# = SelezionaCartella    #09.04# = SalvaFile
'  #90#  LOG_INTERNO = scrittura del log della classe
'        #90.01# = LOG_Scrivi           #90.02# = LOG_Errore       #90.03# = LOG_Passo (solo se LogDettagliato)
'        #90.04# = LOG_Cartella (crea la sottocartella Logs se manca)
'  #91#  SUPPORTO = funzioni private di servizio
'        #91.01# = LeggiAttributi (GetFileAttributesExW)   #91.02# = FileTimeToDate
'        #91.03# = NomeDaFindData                          #91.04# = DimensioneDaParti
'        #91.05# = ScansionaCartella (FindFirstFileW ricorsiva)
'        #91.06# = SvuotaCartella (eliminazione ricorsiva)
'==================================================================================================


'##################################################################################################
'## #01# IMPOSTAZIONI = costanti, strutture e dichiarazioni API 32/64 bit, variabili
'##################################################################################################

'//#01.01# = costanti generali, attributi e codici di errore Windows
Private Const NOME_CLASSE As String = "MSys_TB12_Cls02_FILE"
Private Const MAX_PATH As Long = 260
Private Const PREFISSO_LUNGO As String = "\\?\"
Private Const PREFISSO_UNC As String = "\\?\UNC\"
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400
Private Const MOVEFILE_REPLACE_EXISTING As Long = &H1
Private Const MOVEFILE_COPY_ALLOWED As Long = &H2
Private Const ERROR_ALREADY_EXISTS As Long = 183
Private Const SW_SHOWNORMAL As Long = 1

'//#01.02# = strutture API (nessun puntatore: identiche a 32 e 64 bit)
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FILE_ATTRIBUTE_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName(0 To 519) As Byte
    cAlternateFileName(0 To 27) As Byte
End Type

#If VBA7 Then
    '//#01.03# = dichiarazioni API Unicode per VBA7 (32 e 64 bit): stringhe passate con StrPtr
    Private Declare PtrSafe Function GetFileAttributesExW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal fInfoLevelId As Long, ByRef lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Long
    Private Declare PtrSafe Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwFileAttributes As Long) As Long
    Private Declare PtrSafe Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
    Private Declare PtrSafe Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long
    Private Declare PtrSafe Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, ByVal bFailIfExists As Long) As Long
    Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
    Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As LongPtr
    Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
    Private Declare PtrSafe Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As LongPtr, ByVal lpszShortPath As LongPtr, ByVal cchBuffer As Long) As Long
    Private Declare PtrSafe Function FileTimeToLocalFileTime Lib "kernel32" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
    Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long
    Private Declare PtrSafe Function ShellExecuteW Lib "shell32" (ByVal hwnd As LongPtr, ByVal lpOperation As LongPtr, ByVal lpFile As LongPtr, ByVal lpParameters As LongPtr, ByVal lpDirectory As LongPtr, ByVal nShowCmd As Long) As LongPtr
#Else
    '//#01.04# = dichiarazioni API Unicode per VBA6 (32 bit)
    Private Declare Function GetFileAttributesExW Lib "kernel32" (ByVal lpFileName As Long, ByVal fInfoLevelId As Long, ByRef lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Long
    Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
    Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
    Private Declare Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
    Private Declare Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
    Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
    Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
    Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Private Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
    Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
    Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long
    Private Declare Function ShellExecuteW Lib "shell32" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
#End If

'//#01.05# = variabili della classe
Private m_CartellaLog_s As String
Private m_LogAttivo_b As Boolean
Private m_LogDettagliato_b As Boolean
Private m_UltimoErrore_s As String


'##################################################################################################
'## #02# CICLO_DI_VITA_E_PROPRIETA = inizializzazione, chiusura, impostazioni del log
'##################################################################################################

Private Sub Class_Initialize()
    '//#02.01# = valori predefiniti: log attivo, solo esiti (non ogni passo), cartella <db>\Logs
    m_LogAttivo_b = True
    m_LogDettagliato_b = False
    m_CartellaLog_s = vbNullString
    m_UltimoErrore_s = vbNullString
End Sub

Private Sub Class_Terminate()
    '//#02.02# = nessuna risorsa da rilasciare (gli handle di ricerca si chiudono in #91.05#)
End Sub

'//#02.03# = proprieta
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
'## #03# VERIFICHE = esistenza di file e cartelle, percorso lungo
'##################################################################################################

Public Function FileEsiste(ByVal filePath As String) As Boolean
'//#03.01# = True se il file esiste (anche oltre 260 caratteri)
Const PROC As String = "FileEsiste"
Dim sPos_s As String, uDati As WIN32_FILE_ATTRIBUTE_DATA
On Error GoTo ERRORE
    '//#03.01.01# = percorso vuoto
    sPos_s = "03.01.01"
    If Len(Trim$(filePath)) = 0 Then GoTo USCITA
    '//#03.01.02# = lettura attributi con GetFileAttributesExW
    sPos_s = "03.01.02"
    If Not LeggiAttributi(filePath, uDati) Then GoTo USCITA
    '//#03.01.03# = esito: esiste e non e' una cartella
    sPos_s = "03.01.03"
    FileEsiste = ((uDati.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0)
    LOG_Passo PROC, sPos_s, filePath & " -> " & FileEsiste
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & filePath & "]"
    FileEsiste = False
    Resume USCITA
End Function

Public Function CartellaEsiste(ByVal cartellaPath As String) As Boolean
'//#03.02# = True se la cartella esiste (anche oltre 260 caratteri)
Const PROC As String = "CartellaEsiste"
Dim sPos_s As String, uDati As WIN32_FILE_ATTRIBUTE_DATA
On Error GoTo ERRORE
    '//#03.02.01# = percorso vuoto
    sPos_s = "03.02.01"
    If Len(Trim$(cartellaPath)) = 0 Then GoTo USCITA
    '//#03.02.02# = lettura attributi
    sPos_s = "03.02.02"
    If Not LeggiAttributi(cartellaPath, uDati) Then GoTo USCITA
    '//#03.02.03# = esito: e' una cartella
    sPos_s = "03.02.03"
    CartellaEsiste = ((uDati.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)
    LOG_Passo PROC, sPos_s, cartellaPath & " -> " & CartellaEsiste
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & cartellaPath & "]"
    CartellaEsiste = False
    Resume USCITA
End Function

Public Function IsLongPath(ByVal Percorso As String) As Boolean
'//#03.03# = True se il percorso supera i 260 caratteri (limite delle API classiche)
    IsLongPath = (Len(PercorsoSenzaPrefisso(Percorso)) >= MAX_PATH)
End Function


'##################################################################################################
'## #04# PERCORSI = normalizzazione, prefisso \\?\, percorso breve 8.3, parti del percorso
'##################################################################################################

Public Function NormalizzaPercorso(ByVal Percorso As String) As String
'//#04.01# = percorso per le API Unicode: assoluto con prefisso \\?\ (o \\?\UNC\ per la rete)
Dim s As String
    '//#04.01.01# = pulizia: spazi, barre, doppie barre, barra finale
    s = Trim$(Replace(Percorso, "/", "\"))
    If Len(s) = 0 Then Exit Function
    If Left$(s, 4) = PREFISSO_LUNGO Then NormalizzaPercorso = s: Exit Function
    Do While InStr(3, s, "\\") > 0
        s = Left$(s, 2) & Replace(Mid$(s, 3), "\\", "\")
    Loop
    If Len(s) > 3 And Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    '//#04.01.02# = percorso di rete \\server\cartella -> \\?\UNC\server\cartella
    If Left$(s, 2) = "\\" Then
        NormalizzaPercorso = PREFISSO_UNC & Mid$(s, 3)
    '//#04.01.03# = percorso con unita' C:\... -> \\?\C:\... ; percorso relativo lasciato com'e'
    ElseIf Mid$(s, 2, 1) = ":" Then
        NormalizzaPercorso = PREFISSO_LUNGO & s
    Else
        NormalizzaPercorso = s
    End If
End Function

Public Function PercorsoSenzaPrefisso(ByVal Percorso As String) As String
'//#04.02# = toglie \\?\ e \\?\UNC\ (per mostrare il percorso o passarlo a funzioni classiche)
Dim s As String
    s = Trim$(Percorso)
    If Left$(s, Len(PREFISSO_UNC)) = PREFISSO_UNC Then
        s = "\\" & Mid$(s, Len(PREFISSO_UNC) + 1)
    ElseIf Left$(s, Len(PREFISSO_LUNGO)) = PREFISSO_LUNGO Then
        s = Mid$(s, Len(PREFISSO_LUNGO) + 1)
    End If
    PercorsoSenzaPrefisso = s
End Function

Public Function PercorsoBreve(ByVal Percorso As String) As String
'//#04.03# = percorso breve 8.3 (es. C:\PROGRA~1\...): utile per aprire con la shell oltre 260 caratteri.
'           Se il disco non conserva i nomi brevi restituisce il percorso originale.
Const PROC As String = "PercorsoBreve"
Dim sPos_s As String, sLungo_s As String, sBuf_s As String, nLen_l As Long
On Error GoTo ERRORE
    '//#04.03.01# = chiamata GetShortPathNameW
    sPos_s = "04.03.01"
    sLungo_s = NormalizzaPercorso(Percorso)
    sBuf_s = String$(32767, vbNullChar)
    nLen_l = GetShortPathNameW(StrPtr(sLungo_s), StrPtr(sBuf_s), Len(sBuf_s))
    '//#04.03.02# = risultato senza prefisso
    sPos_s = "04.03.02"
    If nLen_l > 0 And nLen_l < Len(sBuf_s) Then
        PercorsoBreve = PercorsoSenzaPrefisso(Left$(sBuf_s, nLen_l))
    Else
        PercorsoBreve = PercorsoSenzaPrefisso(Percorso)
        LOG_Passo PROC, sPos_s, "nome breve non disponibile (errore Windows " & Err.LastDllError & ")"
    End If
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & Percorso & "]"
    PercorsoBreve = PercorsoSenzaPrefisso(Percorso)
    Resume USCITA
End Function

'//#04.04# = parti del percorso (solo elaborazione del testo: il file puo' anche non esistere)
Public Function GetFileName(ByVal percorsoCompleto As String) As String
Dim s As String, p As Long
    s = PercorsoSenzaPrefisso(Replace(percorsoCompleto, "/", "\"))
    p = InStrRev(s, "\")
    GetFileName = Mid$(s, p + 1)
End Function
Public Function GetFolderPath(ByVal percorsoCompleto As String) As String
Dim s As String, p As Long
    s = PercorsoSenzaPrefisso(Replace(percorsoCompleto, "/", "\"))
    p = InStrRev(s, "\")
    If p > 0 Then GetFolderPath = Left$(s, p - 1)
    If Len(GetFolderPath) = 2 And Right$(GetFolderPath, 1) = ":" Then GetFolderPath = GetFolderPath & "\"
End Function
Public Function GetFileExtension(ByVal percorsoCompleto As String) As String
Dim sNome_s As String, p As Long
    sNome_s = GetFileName(percorsoCompleto)
    p = InStrRev(sNome_s, ".")
    If p > 0 Then GetFileExtension = Mid$(sNome_s, p + 1)
End Function
Public Function UnisciPercorso(ByVal cartella As String, ByVal nome As String) As String
    If Right$(cartella, 1) = "\" Then UnisciPercorso = cartella & nome Else UnisciPercorso = cartella & "\" & nome
End Function


'##################################################################################################
'## #05# CARTELLE = creazione ricorsiva, eliminazione, elenchi
'##################################################################################################

Public Function CreaCartella(ByVal cartellaPath As String) As Boolean
'//#05.01# = crea la cartella e tutte quelle intermedie mancanti (anche oltre 260 caratteri)
Const PROC As String = "CreaCartella"
Dim sPos_s As String, s As String, sRadice_s As String, sResto_s As String
Dim vParti As Variant, i As Long, sCorr_s As String, sLungo_s As String, nErr_l As Long, p As Long
On Error GoTo ERRORE
    '//#05.01.01# = gia' esiste
    sPos_s = "05.01.01"
    If CartellaEsiste(cartellaPath) Then CreaCartella = True: GoTo USCITA
    s = PercorsoSenzaPrefisso(NormalizzaPercorso(cartellaPath))
    If Len(s) = 0 Then GoTo USCITA
    '//#05.01.02# = radice: C:\  oppure  \\server\cartella\
    sPos_s = "05.01.02"
    If Left$(s, 2) = "\\" Then
        p = InStr(3, s, "\"): If p > 0 Then p = InStr(p + 1, s, "\")
        If p = 0 Then GoTo USCITA
        sRadice_s = Left$(s, p): sResto_s = Mid$(s, p + 1)
    ElseIf Mid$(s, 2, 2) = ":\" Then
        sRadice_s = Left$(s, 3): sResto_s = Mid$(s, 4)
    Else
        sRadice_s = vbNullString: sResto_s = s
    End If
    '//#05.01.03# = creazione livello per livello con CreateDirectoryW
    sPos_s = "05.01.03"
    vParti = Split(sResto_s, "\")
    sCorr_s = sRadice_s
    For i = LBound(vParti) To UBound(vParti)
        If Len(vParti(i)) > 0 Then
            sCorr_s = sCorr_s & vParti(i)
            sLungo_s = NormalizzaPercorso(sCorr_s)
            If CreateDirectoryW(StrPtr(sLungo_s), 0) = 0 Then
                nErr_l = Err.LastDllError
                If nErr_l <> ERROR_ALREADY_EXISTS Then
                    LOG_Scrivi PROC, sPos_s, "ERRORE", "livello non creato (errore Windows " & nErr_l & "): " & sCorr_s
                    m_UltimoErrore_s = "#" & sPos_s & "# errore Windows " & nErr_l
                    GoTo USCITA
                End If
            End If
            sCorr_s = sCorr_s & "\"
        End If
    Next i
    '//#05.01.04# = esito
    sPos_s = "05.01.04"
    CreaCartella = CartellaEsiste(cartellaPath)
    LOG_Scrivi PROC, sPos_s, IIf(CreaCartella, "OK", "ERRORE"), "cartella (" & Len(s) & " caratteri): " & s
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & cartellaPath & "]"
    CreaCartella = False
    Resume USCITA
End Function

Public Function EliminaCartella(ByVal cartellaPath As String, Optional ByVal forceDelete As Boolean = False) As Boolean
'//#05.02# = elimina la cartella; con forceDelete=True elimina anche tutto il contenuto
Const PROC As String = "EliminaCartella"
Dim sPos_s As String, sLungo_s As String, nErr_l As Long
On Error GoTo ERRORE
    '//#05.02.01# = controlli
    sPos_s = "05.02.01"
    If Not CartellaEsiste(cartellaPath) Then
        LOG_Scrivi PROC, sPos_s, "AVVISO", "cartella inesistente: " & cartellaPath
        GoTo USCITA
    End If
    sLungo_s = NormalizzaPercorso(cartellaPath)
    '//#05.02.02# = svuotamento ricorsivo (solo se forzata)
    sPos_s = "05.02.02"
    If forceDelete Then SvuotaCartella sLungo_s
    '//#05.02.03# = rimozione con RemoveDirectoryW (fallisce se la cartella non e' vuota)
    sPos_s = "05.02.03"
    If RemoveDirectoryW(StrPtr(sLungo_s)) = 0 Then
        nErr_l = Err.LastDllError
        LOG_Scrivi PROC, sPos_s, "ERRORE", "non eliminata (errore Windows " & nErr_l & IIf(forceDelete, "", ", cartella non vuota?") & "): " & cartellaPath
        m_UltimoErrore_s = "#" & sPos_s & "# errore Windows " & nErr_l
        GoTo USCITA
    End If
    EliminaCartella = True
    LOG_Scrivi PROC, sPos_s, "OK", "eliminata" & IIf(forceDelete, " con il contenuto", "") & ": " & cartellaPath
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & cartellaPath & "]"
    EliminaCartella = False
    Resume USCITA
End Function

Public Function ElencoFile(ByVal cartellaPath As String, Optional ByVal filtro As String = "*", Optional ByVal ricorsivo As Boolean = False) As Collection
'//#05.03# = Collection dei percorsi completi dei file (filtro stile Like, es. "*.mdb")
Const PROC As String = "ElencoFile"
Dim sPos_s As String, coll As New Collection
On Error GoTo ERRORE
    '//#05.03.01# = controlli
    sPos_s = "05.03.01"
    If Not CartellaEsiste(cartellaPath) Then
        LOG_Scrivi PROC, sPos_s, "AVVISO", "cartella inesistente: " & cartellaPath
        GoTo USCITA
    End If
    '//#05.03.02# = scansione con FindFirstFileW
    sPos_s = "05.03.02"
    ScansionaCartella PercorsoSenzaPrefisso(NormalizzaPercorso(cartellaPath)), filtro, ricorsivo, 1, coll
    '//#05.03.03# = esito
    sPos_s = "05.03.03"
    LOG_Passo PROC, sPos_s, coll.Count & " file in " & cartellaPath
USCITA:
    Set ElencoFile = coll
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & cartellaPath & "]"
    Resume USCITA
End Function

Public Function ElencoCartelle(ByVal cartellaPath As String, Optional ByVal filtro As String = "*", Optional ByVal ricorsivo As Boolean = False) As Collection
'//#05.04# = Collection dei percorsi completi delle sottocartelle
Const PROC As String = "ElencoCartelle"
Dim sPos_s As String, coll As New Collection
On Error GoTo ERRORE
    '//#05.04.01# = controlli
    sPos_s = "05.04.01"
    If Not CartellaEsiste(cartellaPath) Then
        LOG_Scrivi PROC, sPos_s, "AVVISO", "cartella inesistente: " & cartellaPath
        GoTo USCITA
    End If
    '//#05.04.02# = scansione
    sPos_s = "05.04.02"
    ScansionaCartella PercorsoSenzaPrefisso(NormalizzaPercorso(cartellaPath)), filtro, ricorsivo, 2, coll
    '//#05.04.03# = esito
    sPos_s = "05.04.03"
    LOG_Passo PROC, sPos_s, coll.Count & " cartelle in " & cartellaPath
USCITA:
    Set ElencoCartelle = coll
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & cartellaPath & "]"
    Resume USCITA
End Function


'##################################################################################################
'## #06# FILE = copia, spostamento, eliminazione
'##################################################################################################

Public Function CopiaFile(ByVal sourcePath As String, ByVal destinationPath As String, Optional ByVal overwrite As Boolean = True) As Boolean
'//#06.01# = copia un file; se la destinazione e' una cartella mantiene il nome; crea le cartelle mancanti
Const PROC As String = "CopiaFile"
Dim sPos_s As String, sDest_s As String, nErr_l As Long
On Error GoTo ERRORE
    '//#06.01.01# = controlli
    sPos_s = "06.01.01"
    If Not FileEsiste(sourcePath) Then
        LOG_Scrivi PROC, sPos_s, "ERRORE", "file di origine inesistente: " & sourcePath
        m_UltimoErrore_s = "#" & sPos_s & "# origine inesistente"
        GoTo USCITA
    End If
    '//#06.01.02# = destinazione: cartella esistente -> stesso nome; cartella mancante -> creata
    sPos_s = "06.01.02"
    sDest_s = PercorsoSenzaPrefisso(destinationPath)
    If CartellaEsiste(sDest_s) Then sDest_s = UnisciPercorso(sDest_s, GetFileName(sourcePath))
    If Not CartellaEsiste(GetFolderPath(sDest_s)) Then CreaCartella GetFolderPath(sDest_s)
    '//#06.01.03# = CopyFileW
    sPos_s = "06.01.03"
    If CopyFileW(StrPtr(NormalizzaPercorso(sourcePath)), StrPtr(NormalizzaPercorso(sDest_s)), IIf(overwrite, 0, 1)) = 0 Then
        nErr_l = Err.LastDllError
        LOG_Scrivi PROC, sPos_s, "ERRORE", "copia non riuscita (errore Windows " & nErr_l & "): " & sourcePath & " -> " & sDest_s
        m_UltimoErrore_s = "#" & sPos_s & "# errore Windows " & nErr_l
        GoTo USCITA
    End If
    CopiaFile = True
    LOG_Scrivi PROC, sPos_s, "OK", sourcePath & " -> " & sDest_s
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & sourcePath & "]"
    CopiaFile = False
    Resume USCITA
End Function

Public Function SpostaFile(ByVal sourcePath As String, ByVal destinationPath As String, Optional ByVal overwrite As Boolean = False) As Boolean
'//#06.02# = sposta o rinomina un file (anche tra dischi diversi)
Const PROC As String = "SpostaFile"
Dim sPos_s As String, sDest_s As String, nErr_l As Long, nFlag_l As Long
On Error GoTo ERRORE
    '//#06.02.01# = controlli
    sPos_s = "06.02.01"
    If Not FileEsiste(sourcePath) Then
        LOG_Scrivi PROC, sPos_s, "ERRORE", "file di origine inesistente: " & sourcePath
        m_UltimoErrore_s = "#" & sPos_s & "# origine inesistente"
        GoTo USCITA
    End If
    '//#06.02.02# = destinazione
    sPos_s = "06.02.02"
    sDest_s = PercorsoSenzaPrefisso(destinationPath)
    If CartellaEsiste(sDest_s) Then sDest_s = UnisciPercorso(sDest_s, GetFileName(sourcePath))
    If Not CartellaEsiste(GetFolderPath(sDest_s)) Then CreaCartella GetFolderPath(sDest_s)
    '//#06.02.03# = MoveFileExW
    sPos_s = "06.02.03"
    nFlag_l = MOVEFILE_COPY_ALLOWED
    If overwrite Then nFlag_l = nFlag_l Or MOVEFILE_REPLACE_EXISTING
    If MoveFileExW(StrPtr(NormalizzaPercorso(sourcePath)), StrPtr(NormalizzaPercorso(sDest_s)), nFlag_l) = 0 Then
        nErr_l = Err.LastDllError
        LOG_Scrivi PROC, sPos_s, "ERRORE", "spostamento non riuscito (errore Windows " & nErr_l & "): " & sourcePath & " -> " & sDest_s
        m_UltimoErrore_s = "#" & sPos_s & "# errore Windows " & nErr_l
        GoTo USCITA
    End If
    SpostaFile = True
    LOG_Scrivi PROC, sPos_s, "OK", sourcePath & " -> " & sDest_s
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & sourcePath & "]"
    SpostaFile = False
    Resume USCITA
End Function

Public Function EliminaFile(ByVal filePath As String) As Boolean
'//#06.03# = elimina un file, anche se in sola lettura
Const PROC As String = "EliminaFile"
Dim sPos_s As String, sLungo_s As String, nErr_l As Long
On Error GoTo ERRORE
    '//#06.03.01# = controlli
    sPos_s = "06.03.01"
    If Not FileEsiste(filePath) Then
        LOG_Scrivi PROC, sPos_s, "AVVISO", "file inesistente: " & filePath
        GoTo USCITA
    End If
    '//#06.03.02# = tolgo la sola lettura
    sPos_s = "06.03.02"
    sLungo_s = NormalizzaPercorso(filePath)
    SetFileAttributesW StrPtr(sLungo_s), FILE_ATTRIBUTE_NORMAL
    '//#06.03.03# = DeleteFileW
    sPos_s = "06.03.03"
    If DeleteFileW(StrPtr(sLungo_s)) = 0 Then
        nErr_l = Err.LastDllError
        LOG_Scrivi PROC, sPos_s, "ERRORE", "non eliminato (errore Windows " & nErr_l & "): " & filePath
        m_UltimoErrore_s = "#" & sPos_s & "# errore Windows " & nErr_l
        GoTo USCITA
    End If
    EliminaFile = True
    LOG_Scrivi PROC, sPos_s, "OK", "eliminato: " & filePath
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & filePath & "]"
    EliminaFile = False
    Resume USCITA
End Function


'##################################################################################################
'## #07# ATTRIBUTI_E_PROPRIETA = attributi, dimensione, date, scheda proprieta'
'##################################################################################################

'//#07.01# = attributi: valore numerico (-1 se il file non esiste) e descrizione in italiano
Public Function AttributiFile(ByVal Percorso As String) As Long
Dim uDati As WIN32_FILE_ATTRIBUTE_DATA
    If LeggiAttributi(Percorso, uDati) Then AttributiFile = uDati.dwFileAttributes Else AttributiFile = -1
End Function
Public Function AttributiTesto(ByVal Percorso As String) As String
Dim a As Long, s As String
    a = AttributiFile(Percorso)
    If a = -1 Then AttributiTesto = "Non disponibile": Exit Function
    If a And &H1 Then s = s & "Sola lettura, "
    If a And &H2 Then s = s & "Nascosto, "
    If a And &H4 Then s = s & "Sistema, "
    If a And &H10 Then s = s & "Cartella, "
    If a And &H20 Then s = s & "Archivio, "
    If a And &H100 Then s = s & "Temporaneo, "
    If a And &H400 Then s = s & "Collegamento, "
    If a And &H800 Then s = s & "Compresso, "
    If a And &H1000 Then s = s & "Offline, "
    If a And &H2000 Then s = s & "Non indicizzato, "
    If a And &H4000 Then s = s & "Crittografato, "
    If Len(s) = 0 Then AttributiTesto = "Normale" Else AttributiTesto = Left$(s, Len(s) - 2)
End Function

'//#07.02# = dimensione in byte (Double: oltre 2 GB) e in formato leggibile
Public Function DimensioneFile(ByVal filePath As String) As Double
Dim uDati As WIN32_FILE_ATTRIBUTE_DATA
    If LeggiAttributi(filePath, uDati) Then DimensioneFile = DimensioneDaParti(uDati.nFileSizeHigh, uDati.nFileSizeLow) Else DimensioneFile = -1
End Function
Public Function GetFileSizeFormatted(ByVal filePath As String) As String
Dim d As Double
    d = DimensioneFile(filePath)
    Select Case d
        Case Is < 0:          GetFileSizeFormatted = vbNullString
        Case Is < 1024:       GetFileSizeFormatted = Format$(d, "0") & " bytes"
        Case Is < 1048576:    GetFileSizeFormatted = Format$(d / 1024, "0.0") & " KB"
        Case Is < 1073741824: GetFileSizeFormatted = Format$(d / 1048576, "0.0") & " MB"
        Case Else:            GetFileSizeFormatted = Format$(d / 1073741824, "0.0") & " GB"
    End Select
End Function

'//#07.03# = date in ora locale (0 se il file non esiste)
Public Function DataCreazione(ByVal Percorso As String) As Date
Dim uDati As WIN32_FILE_ATTRIBUTE_DATA
    If LeggiAttributi(Percorso, uDati) Then DataCreazione = FileTimeToDate(uDati.ftCreationTime)
End Function
Public Function DataModifica(ByVal Percorso As String) As Date
Dim uDati As WIN32_FILE_ATTRIBUTE_DATA
    If LeggiAttributi(Percorso, uDati) Then DataModifica = FileTimeToDate(uDati.ftLastWriteTime)
End Function
Public Function DataUltimoAccesso(ByVal Percorso As String) As Date
Dim uDati As WIN32_FILE_ATTRIBUTE_DATA
    If LeggiAttributi(Percorso, uDati) Then DataUltimoAccesso = FileTimeToDate(uDati.ftLastAccessTime)
End Function

Public Function ProprietaFile(ByVal filePath As String) As String
'//#07.04# = scheda proprieta' in testo (anche per percorsi oltre 260 caratteri)
Const PROC As String = "ProprietaFile"
Dim sPos_s As String, uDati As WIN32_FILE_ATTRIBUTE_DATA, s As String, sPerc_s As String
On Error GoTo ERRORE
    '//#07.04.01# = lettura attributi e date
    sPos_s = "07.04.01"
    If Not LeggiAttributi(filePath, uDati) Then
        ProprietaFile = "File non trovato: " & filePath
        GoTo USCITA
    End If
    '//#07.04.02# = composizione del testo
    sPos_s = "07.04.02"
    sPerc_s = PercorsoSenzaPrefisso(filePath)
    s = "Percorso: " & sPerc_s & vbCrLf
    s = s & "Nome: " & GetFileName(sPerc_s) & vbCrLf
    s = s & "Cartella: " & GetFolderPath(sPerc_s) & vbCrLf
    s = s & "Estensione: " & GetFileExtension(sPerc_s) & vbCrLf
    s = s & "Dimensione: " & Format$(DimensioneDaParti(uDati.nFileSizeHigh, uDati.nFileSizeLow), "#,##0") & " byte (" & GetFileSizeFormatted(filePath) & ")" & vbCrLf
    s = s & "Creazione: " & Format$(FileTimeToDate(uDati.ftCreationTime), "dd/mm/yyyy hh:nn:ss") & vbCrLf
    s = s & "Ultima modifica: " & Format$(FileTimeToDate(uDati.ftLastWriteTime), "dd/mm/yyyy hh:nn:ss") & vbCrLf
    s = s & "Ultimo accesso: " & Format$(FileTimeToDate(uDati.ftLastAccessTime), "dd/mm/yyyy hh:nn:ss") & vbCrLf
    s = s & "Attributi: " & AttributiTesto(filePath) & vbCrLf
    s = s & "Lunghezza percorso: " & Len(sPerc_s) & " caratteri" & IIf(IsLongPath(sPerc_s), " (percorso lungo)", "")
    ProprietaFile = s
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & filePath & "]"
    ProprietaFile = "Errore: " & Err.Description
    Resume USCITA
End Function

Public Sub MostraProprietaFile(ByVal filePath As String)
'//#07.05# = mostra la scheda proprieta' in una finestra di messaggio
    MsgBox ProprietaFile(filePath), vbInformation, "Informazioni sul file"
    LOG_Scrivi "MostraProprietaFile", "07.05", "INFO", "proprieta' mostrate: " & filePath
End Sub


'##################################################################################################
'## #08# APERTURA = apertura con l'applicazione predefinita (anche oltre 260 caratteri)
'##################################################################################################

Public Function ApriFile(ByVal percorsoCompleto As String, Optional ByVal showCmd As Long = SW_SHOWNORMAL) As Boolean
'//#08.01# = apre file o cartella con l'applicazione associata (ShellExecuteW)
Const PROC As String = "ApriFile"
Dim sPos_s As String, sShell_s As String, sOper_s As String
#If VBA7 Then
    Dim r As LongPtr
#Else
    Dim r As Long
#End If
On Error GoTo ERRORE
    '//#08.01.01# = controlli
    sPos_s = "08.01.01"
    If Not FileEsiste(percorsoCompleto) And Not CartellaEsiste(percorsoCompleto) Then
        LOG_Scrivi PROC, sPos_s, "ERRORE", "percorso inesistente: " & percorsoCompleto
        m_UltimoErrore_s = "#" & sPos_s & "# percorso inesistente"
        GoTo USCITA
    End If
    '//#08.01.02# = percorso per la shell: oltre 260 caratteri uso il nome breve 8.3
    sPos_s = "08.01.02"
    sShell_s = PercorsoSenzaPrefisso(percorsoCompleto)
    If IsLongPath(sShell_s) Then sShell_s = PercorsoBreve(sShell_s)
    If CartellaEsiste(percorsoCompleto) Then sOper_s = "explore" Else sOper_s = "open"
    '//#08.01.03# = ShellExecuteW (successo se il valore restituito e' maggiore di 32)
    sPos_s = "08.01.03"
    r = ShellExecuteW(0, StrPtr(sOper_s), StrPtr(sShell_s), 0, 0, showCmd)
    ApriFile = (r > 32)
    LOG_Scrivi PROC, sPos_s, IIf(ApriFile, "OK", "ERRORE"), sOper_s & " (codice " & CLng(r) & "): " & sShell_s
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & percorsoCompleto & "]"
    ApriFile = False
    Resume USCITA
End Function

Public Function ApriCartella(ByVal cartellaPath As String) As Boolean
'//#08.02# = apre la cartella in Esplora file (anche oltre 260 caratteri)
    ApriCartella = ApriFile(cartellaPath)
End Function


'##################################################################################################
'## #09# FINESTRE_DI_DIALOGO = selezione file, file multipli, cartella, salvataggio
'##        Application.FileDialog: 2 = Salva con nome, 3 = Selezione file, 4 = Selezione cartella
'##################################################################################################

Public Function SelezionaFile(Optional ByVal filtroDescrizione As String = "Tutti i file (*.*)", _
                              Optional ByVal filtroEstensione As String = "*.*", _
                              Optional ByVal cartellaIniziale As String = vbNullString, _
                              Optional ByVal titolo As String = "Seleziona un file") As String
'//#09.01# = un file; stringa vuota se annullato
Dim coll As Collection
    Set coll = DIALOGO_File(False, filtroDescrizione, filtroEstensione, cartellaIniziale, titolo)
    If coll.Count > 0 Then SelezionaFile = coll(1)
End Function

Public Function SelezionaFileMultipli(Optional ByVal filtroDescrizione As String = "Tutti i file (*.*)", _
                                      Optional ByVal filtroEstensione As String = "*.*", _
                                      Optional ByVal cartellaIniziale As String = vbNullString, _
                                      Optional ByVal titolo As String = "Seleziona uno o piu' file") As Collection
'//#09.02# = piu' file; Collection vuota se annullato
    Set SelezionaFileMultipli = DIALOGO_File(True, filtroDescrizione, filtroEstensione, cartellaIniziale, titolo)
End Function

Public Function SelezionaCartella(Optional ByVal cartellaIniziale As String = vbNullString, _
                                  Optional ByVal titolo As String = "Seleziona una cartella") As String
'//#09.03# = una cartella; stringa vuota se annullato
Const PROC As String = "SelezionaCartella"
Dim fd As Object
On Error GoTo ERRORE
    Set fd = Application.fileDialog(4)
    fd.Title = titolo
    If Len(cartellaIniziale) > 0 Then fd.InitialFileName = UnisciPercorso(PercorsoSenzaPrefisso(cartellaIniziale), vbNullString)
    If fd.Show = -1 Then SelezionaCartella = fd.SelectedItems(1)
    LOG_Passo PROC, "09.03", "scelta: " & SelezionaCartella
USCITA:
    Set fd = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, "09.03", Err.Number, Err.Description
    Resume USCITA
End Function

Public Function SalvaFile(Optional ByVal filtroDescrizione As String = "Tutti i file (*.*)", _
                          Optional ByVal filtroEstensione As String = "*.*", _
                          Optional ByVal cartellaIniziale As String = vbNullString, _
                          Optional ByVal nomeFileSuggerito As String = vbNullString, _
                          Optional ByVal titolo As String = "Salva file") As String
'//#09.04# = percorso scelto per salvare (la finestra Salva con nome di Access non usa filtri)
Const PROC As String = "SalvaFile"
Dim fd As Object, sIniz_s As String
On Error GoTo ERRORE
    Set fd = Application.fileDialog(2)
    fd.Title = titolo
    If Len(cartellaIniziale) > 0 Then sIniz_s = UnisciPercorso(PercorsoSenzaPrefisso(cartellaIniziale), vbNullString)
    fd.InitialFileName = sIniz_s & nomeFileSuggerito
    If fd.Show = -1 Then SalvaFile = fd.SelectedItems(1)
    LOG_Passo PROC, "09.04", "scelta: " & SalvaFile
USCITA:
    Set fd = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, "09.04", Err.Number, Err.Description
    Resume USCITA
End Function

Private Function DIALOGO_File(ByVal bMulti As Boolean, ByVal sDescr As String, ByVal sEst As String, ByVal sIniz As String, ByVal sTitolo As String) As Collection
'//#09.01# / #09.02# = finestra di selezione file comune
Const PROC As String = "SelezionaFile"
Dim fd As Object, coll As New Collection, i As Long
On Error GoTo ERRORE
    Set fd = Application.fileDialog(3)
    fd.allowMultiSelect = bMulti
    fd.Title = sTitolo
    fd.Filters.Clear
    fd.Filters.Add sDescr, sEst
    If Len(sIniz) > 0 Then fd.InitialFileName = UnisciPercorso(PercorsoSenzaPrefisso(sIniz), vbNullString)
    If fd.Show = -1 Then
        For i = 1 To fd.SelectedItems.Count
            coll.Add fd.SelectedItems(i)
        Next i
    End If
    LOG_Passo PROC, IIf(bMulti, "09.02", "09.01"), coll.Count & " file scelti"
USCITA:
    Set DIALOGO_File = coll
    Set fd = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, IIf(bMulti, "09.02", "09.01"), Err.Number, Err.Description
    Resume USCITA
End Function


'##################################################################################################
'## #90# LOG_INTERNO = scrittura del log della classe in <db>\Logs\MSys_TB12_Cls02_FILE_AAAA_MM_GG.txt
'##################################################################################################

Private Sub LOG_Scrivi(ByVal sFunz As String, ByVal sPos As String, ByVal sEsito As String, ByVal sTesto As String)
'//#90.01# = la stessa riga in Debug.Print e nel file; il log non deve mai bloccare la classe
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
'//#90.02# = errore con la posizione esatta nel codice; resta disponibile in UltimoErrore
    m_UltimoErrore_s = "#" & sPos & "# " & sFunz & " - errore " & nErr & ": " & sDescr
    LOG_Scrivi sFunz, sPos, "ERRORE", "errore " & nErr & ": " & sDescr & " -> codice fermo a #" & sPos & "#"
End Sub

Private Sub LOG_Passo(ByVal sFunz As String, ByVal sPos As String, ByVal sTesto As String)
'//#90.03# = passo intermedio, scritto solo con LogDettagliato = True
    If m_LogDettagliato_b Then LOG_Scrivi sFunz, sPos, "INFO", sTesto
End Sub

Private Function LOG_Cartella() As String
'//#90.04# = cartella del log: impostata dall'utente oppure <cartella del database>\Logs (creata se manca)
    On Error Resume Next
    If Len(m_CartellaLog_s) > 0 Then LOG_Cartella = m_CartellaLog_s Else LOG_Cartella = CurrentProject.path & "\Logs"
    If Len(Dir$(LOG_Cartella, vbDirectory)) = 0 Then MkDir LOG_Cartella
End Function


'##################################################################################################
'## #91# SUPPORTO = funzioni private di servizio
'##################################################################################################

Private Function LeggiAttributi(ByVal sPercorso As String, ByRef uDati As WIN32_FILE_ATTRIBUTE_DATA) As Boolean
'//#91.01# = GetFileAttributesExW con prefisso \\?\: True se il percorso esiste
Dim sLungo_s As String
    sLungo_s = NormalizzaPercorso(sPercorso)
    If Len(sLungo_s) = 0 Then Exit Function
    LeggiAttributi = (GetFileAttributesExW(StrPtr(sLungo_s), 0, uDati) <> 0)
End Function

Private Function FileTimeToDate(ByRef ft As FILETIME) As Date
'//#91.02# = FILETIME (UTC) -> data e ora locale
Dim ftL As FILETIME, st As SYSTEMTIME
    If ft.dwLowDateTime = 0 And ft.dwHighDateTime = 0 Then Exit Function
    If FileTimeToLocalFileTime(ft, ftL) = 0 Then Exit Function
    If FileTimeToSystemTime(ftL, st) = 0 Then Exit Function
    FileTimeToDate = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

Private Function NomeDaFindData(ByRef fd As WIN32_FIND_DATA) As String
'//#91.03# = nome del file dal buffer Unicode di WIN32_FIND_DATA
Dim i As Long, c As Long, s As String
    For i = 0 To 518 Step 2
        c = fd.cFileName(i) + fd.cFileName(i + 1) * 256&
        If c = 0 Then Exit For
        s = s & ChrW$(c)
    Next i
    NomeDaFindData = s
End Function

Private Function DimensioneDaParti(ByVal nAlta As Long, ByVal nBassa As Long) As Double
'//#91.04# = dimensione a 64 bit da due Long senza segno
Dim dBassa As Double
    dBassa = nBassa
    If dBassa < 0 Then dBassa = dBassa + 4294967296#
    DimensioneDaParti = CDbl(nAlta) * 4294967296# + dBassa
End Function

Private Sub ScansionaCartella(ByVal sCartella As String, ByVal sFiltro As String, ByVal bRicorsivo As Boolean, ByVal iModo As Integer, ByRef coll As Collection)
'//#91.05# = FindFirstFileW/FindNextFileW; iModo 1 = file, 2 = cartelle; i collegamenti (reparse) non si seguono
Dim fd As WIN32_FIND_DATA, sNome_s As String, sPieno_s As String, bCart As Boolean
#If VBA7 Then
    Dim h As LongPtr
#Else
    Dim h As Long
#End If
    h = FindFirstFileW(StrPtr(NormalizzaPercorso(sCartella & "\*")), fd)
    If h = -1 Then Exit Sub
    Do
        sNome_s = NomeDaFindData(fd)
        If sNome_s <> "." And sNome_s <> ".." Then
            sPieno_s = sCartella & "\" & sNome_s
            bCart = ((fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)
            If bCart Then
                If iModo = 2 And (sNome_s Like sFiltro) Then coll.Add sPieno_s
                If bRicorsivo And (fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) = 0 Then ScansionaCartella sPieno_s, sFiltro, bRicorsivo, iModo, coll
            Else
                If iModo = 1 And (sNome_s Like sFiltro) Then coll.Add sPieno_s
            End If
        End If
    Loop While FindNextFileW(h, fd) <> 0
    FindClose h
End Sub

Private Sub SvuotaCartella(ByVal sCartellaLunga As String)
'//#91.06# = elimina ricorsivamente file e sottocartelle (percorso gia' con prefisso \\?\)
Dim fd As WIN32_FIND_DATA, sNome_s As String, sPieno_s As String
#If VBA7 Then
    Dim h As LongPtr
#Else
    Dim h As Long
#End If
    h = FindFirstFileW(StrPtr(sCartellaLunga & "\*"), fd)
    If h = -1 Then Exit Sub
    Do
        sNome_s = NomeDaFindData(fd)
        If sNome_s <> "." And sNome_s <> ".." Then
            sPieno_s = sCartellaLunga & "\" & sNome_s
            If (fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
                If (fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) = 0 Then SvuotaCartella sPieno_s
                SetFileAttributesW StrPtr(sPieno_s), FILE_ATTRIBUTE_NORMAL
                RemoveDirectoryW StrPtr(sPieno_s)
            Else
                SetFileAttributesW StrPtr(sPieno_s), FILE_ATTRIBUTE_NORMAL
                DeleteFileW StrPtr(sPieno_s)
            End If
        End If
    Loop While FindNextFileW(h, fd) <> 0
    FindClose h
End Sub
