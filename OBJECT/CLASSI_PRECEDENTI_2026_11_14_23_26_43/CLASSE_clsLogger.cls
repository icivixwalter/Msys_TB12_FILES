VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSE_clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' =====================================================
' CLASSE: clsLogger
' RESPONSABILITA': Gestione centralizzata dei log di sistema
' CREAZIONE: Estratta da CLASSE_ATTI_TB01_GESTIONE
' =====================================================
Option Compare Database
Option Explicit

' Metodo principale per la scrittura dei log
' Parametri:
'   nomeFunzione      - Nome della funzione/procedura che ha generato l'errore
'   NumeroErrore      - Numero dell'errore (Err.Number)
'   DescrizioneErrore - Descrizione dell'errore (Err.Description)
'   RigaErrore        - Numero di riga dove si è verificato l'errore (Erl)
'
' Output: Scrive una riga di log nel file Logs\LogErrori.txt
'
Public Sub ScriviLogErrore(nomeFunzione As String, NumeroErrore As Long, DescrizioneErrore As String, RigaErrore As Long)
    On Error GoTo GestioneErrore

    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim testo As String
    Dim dataOra As String
    Dim cartellaLogs As String

    ' Format data/ora per log
    dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")

    ' Imposta la cartella Logs nel percorso del database
    cartellaLogs = CurrentProject.path & "\Logs"

    ' Crea la cartella Logs se non esiste
    If Dir(cartellaLogs, vbDirectory) = "" Then
        MkDir cartellaLogs
    End If

    ' Costruisci il percorso completo del file di log
    percorsoLog = cartellaLogs & "\LogErrori.txt"

    ' Componi il testo da scrivere
    testo = dataOra & " | Classe: CLASSE_clsLogger | Funzione: " & nomeFunzione & _
            " | Errore #: " & NumeroErrore & " | Descrizione: " & DescrizioneErrore & _
            " | Riga: " & RigaErrore

    ' Scrive nel file in modalità Append
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        Print #fileNum, testo
    Close #fileNum

    Exit Sub

GestioneErrore:
    ' Gestione errori interni alla funzione di log
    MsgBox "Errore durante la scrittura del log." & vbCrLf & _
           "Descrizione: " & Err.Description & vbCrLf & _
           "Errore #: " & Err.Number, vbCritical, "Errore ScriviLogErrore"

    ' Scrittura di emergenza in un file alternativo
    On Error Resume Next
    percorsoLog = CurrentProject.path & "\LogErrore_di_Sistema.txt"
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        Print #fileNum, Format(Now, "yyyy-mm-dd HH:nn:ss") & _
              " | Scrittura log fallita nella funzione: ScriviLogErrore" & _
              " | Errore interno #: " & Err.Number & " - " & Err.Description
    Close #fileNum
End Sub

' Metodo opzionale per scrivere messaggi informativi (non solo errori)
' Parametri:
'   nomeFunzione      - Nome della funzione/procedura
'   messaggio         - Messaggio informativo da loggare
'
Public Sub ScriviInfo(nomeFunzione As String, messaggio As String)
    On Error GoTo GestioneErrore

    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim dataOra As String
    Dim cartellaLogs As String

    ' Format data/ora per log
    dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")

    ' Imposta la cartella Logs nel percorso del database
    cartellaLogs = CurrentProject.path & "\Logs"

    ' Crea la cartella Logs se non esiste
    If Dir(cartellaLogs, vbDirectory) = "" Then
        MkDir cartellaLogs
    End If

    ' Costruisci il percorso completo del file di log (puoi usare un file separato per gli info)
    percorsoLog = cartellaLogs & "\LogInfo.txt"

    ' Scrive nel file in modalità Append
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        Print #fileNum, dataOra & " | " & nomeFunzione & " | INFO: " & messaggio
    Close #fileNum

    Exit Sub

GestioneErrore:
    ' Se fallisce anche il log info, silenziosamente ignora
End Sub

