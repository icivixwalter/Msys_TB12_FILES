VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Msys_TB12_FILES_CLASSE_clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ====================================================================================================================
' CLASSE: My_Msys_TB12_FILES_CLASSE_clsLogger
' RESPONSABILITA': Gestione centralizzata dei log di sistema
' CREAZIONE: Estratta da CLASSE_ATTI_TB01_GESTIONE
' VERSIONE: 1.0
' DATA: 2026-06-18
' ====================================================================================================================
'
' ====================================================================================================================
' INDICE DELLE ATTIVITA'
' ====================================================================================================================
' 01_ATTIVITA = Inizializzazione e configurazione logger
' 02_ATTIVITA = Creazione directory Logs se non esiste
' 03_ATTIVITA = Scrittura log di errore
' 04_ATTIVITA = Scrittura log informativo
' 05_ATTIVITA = Scrittura log evento (alias)
' 06_ATTIVITA = Scrittura log su file personalizzato
' 07_ATTIVITA = Gestione errori interni del logger
' 08_ATTIVITA = Scrittura log di emergenza
' 09_ATTIVITA = Utility per pulizia log
' 10_ATTIVITA = Utility per dimensione log
' 11_ATTIVITA = Utility per conteggio log
' 12_ATTIVITA = Utility per visualizzare log
' ====================================================================================================================

Option Compare Database
Option Explicit


' ********************************************************************************************************************
' *** INIZIO 01_ATTIVITA = Inizializzazione e configurazione logger ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: Class_Initialize
' Descrizione: Inizializza le variabili del logger
' ====================================================================================================================
Private Sub Class_Initialize()
    ' 01_01 = Inizializzazione logger
    ' Il logger è pronto per l'uso
End Sub

' ====================================================================================================================
' Metodo: Class_Terminate
' Descrizione: Pulizia risorse al termine
' ====================================================================================================================
Private Sub Class_Terminate()
    ' 01_02 = Terminazione logger
    ' Rilascio risorse se necessario
End Sub

' ********************************************************************************************************************
' *** FINE 01_ATTIVITA = Inizializzazione e configurazione logger ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 02_ATTIVITA = Creazione directory Logs se non esiste ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: CreaDirectoryLogs
' Descrizione: Verifica e crea la cartella Logs se assente
' Parametri: Nessuno
' Output: Percorso della cartella Logs
' ====================================================================================================================
Private Function CreaDirectoryLogs() As String
    ' 02_01 = Verifica esistenza cartella Logs
    Dim cartellaLogs As String
    
    ' 02_02 = Imposta percorso cartella Logs
    cartellaLogs = CurrentProject.path & "\Logs"
    
    ' 02_03 = Controlla se la cartella esiste
    If Dir(cartellaLogs, vbDirectory) = "" Then
        ' 02_04 = Crea la cartella Logs
        MkDir cartellaLogs
    End If
    
    ' 02_05 = Restituisce il percorso
    CreaDirectoryLogs = cartellaLogs
End Function

' ********************************************************************************************************************
' *** FINE 02_ATTIVITA = Creazione directory Logs se non esiste ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 03_ATTIVITA = Scrittura log di errore ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: ScriviLogErrore
' Descrizione: Scrive un log di errore nel file LogErrori.txt
' Parametri:
'   nomeFunzione      - Nome della funzione/procedura che ha generato l'errore
'   NumeroErrore      - Numero dell'errore (Err.Number)
'   DescrizioneErrore - Descrizione dell'errore (Err.Description)
'   RigaErrore        - Numero di riga dove si è verificato l'errore (Erl)
' Output: Scrive una riga di log nel file Logs\LogErrori.txt
' ====================================================================================================================
Public Sub ScriviLogErrore(nomeFunzione As String, NumeroErrore As Long, DescrizioneErrore As String, RigaErrore As Long)
    ' 03_01 = Gestione errori interna
    On Error GoTo GestioneErrore

    ' 03_02 = Dichiarazione variabili
    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim testo As String
    Dim dataOra As String
    Dim cartellaLogs As String

    ' 03_03 = Formatta data/ora per log
    dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")

    ' 03_04 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()

    ' 03_05 = Costruisci percorso completo file di log
    percorsoLog = cartellaLogs & "\LogErrori.txt"

    ' 03_06 = Componi il testo da scrivere
    testo = dataOra & " | Classe: My_Msys_TB12_FILES_CLASSE_clsLogger | Funzione: " & nomeFunzione & _
            " | Errore #: " & NumeroErrore & " | Descrizione: " & DescrizioneErrore & _
            " | Riga: " & RigaErrore

    ' 03_07 = Scrive nel file in modalità Append
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        Print #fileNum, testo
    Close #fileNum

    ' 03_08 = Uscita normale
    Exit Sub

' 03_09 = Gestione errori interni del logger
GestioneErrore:
    ' 03_10 = Gestione errori interni alla funzione di log
    MsgBox "Errore durante la scrittura del log." & vbCrLf & _
           "Descrizione: " & Err.Description & vbCrLf & _
           "Errore #: " & Err.Number, vbCritical, "Errore ScriviLogErrore"
    
    ' 03_11 = Chiamata scrittura emergenza
    ScriviLogEmergenza "ScriviLogErrore", Err.Number, Err.Description
End Sub

' ********************************************************************************************************************
' *** FINE 03_ATTIVITA = Scrittura log di errore ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 04_ATTIVITA = Scrittura log informativo ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: ScriviInfo
' Descrizione: Scrive un messaggio informativo nel file LogInfo.txt
' Parametri:
'   nomeFunzione      - Nome della funzione/procedura
'   messaggio         - Messaggio informativo da loggare
' Output: Scrive una riga di log nel file Logs\LogInfo.txt
' ====================================================================================================================
Public Sub ScriviInfo(nomeFunzione As String, messaggio As String)
    ' 04_01 = Gestione errori interna (silenziosa)
    On Error GoTo GestioneErrore

    ' 04_02 = Dichiarazione variabili
    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim dataOra As String
    Dim cartellaLogs As String

    ' 04_03 = Formatta data/ora per log
    dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")

    ' 04_04 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()

    ' 04_05 = Costruisci percorso completo file di log
    percorsoLog = cartellaLogs & "\LogInfo.txt"

    ' 04_06 = Scrive nel file in modalità Append
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        ' 04_07 = Componi e scrivi il testo
        Print #fileNum, dataOra & " | " & nomeFunzione & " | INFO: " & messaggio
    Close #fileNum

    ' 04_08 = Uscita normale
    Exit Sub

' 04_09 = Gestione errori interni (silenziosa)
GestioneErrore:
    ' 04_10 = Se fallisce, silenziosamente ignora
    ' Non fare nulla per non interrompere il flusso principale
End Sub

' ********************************************************************************************************************
' *** FINE 04_ATTIVITA = Scrittura log informativo ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 05_ATTIVITA = Scrittura log evento (alias) ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: ScriviLogEvento
' Descrizione: Metodo alias di ScriviInfo per compatibilità
' Parametri:
'   nomeFunzione      - Nome della funzione/procedura
'   messaggio         - Messaggio evento da loggare
' Output: Scrive una riga di log nel file Logs\LogInfo.txt
' ====================================================================================================================
Public Sub ScriviLogEvento(nomeFunzione As String, messaggio As String)
    ' 05_01 = Chiamata al metodo ScriviInfo
    ' 05_02 = Mantiene compatibilità con codice esistente
    ScriviInfo nomeFunzione, messaggio
End Sub

' ********************************************************************************************************************
' *** FINE 05_ATTIVITA = Scrittura log evento (alias) ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 06_ATTIVITA = Scrittura log su file personalizzato ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: ScriviLogPersonalizzato
' Descrizione: Scrive un log su un file specificato dall'utente
' Parametri:
'   nomeFunzione      - Nome della funzione/procedura
'   messaggio         - Messaggio da loggare
'   nomeFile          - Nome del file di log (es. "LogOperazioni.txt")
' Output: Scrive una riga di log nel file Logs\[nomefile].txt
' ====================================================================================================================
Public Sub ScriviLogPersonalizzato(nomeFunzione As String, messaggio As String, nomeFile As String)
    ' 06_01 = Gestione errori interna (silenziosa)
    On Error GoTo GestioneErrore

    ' 06_02 = Dichiarazione variabili
    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim dataOra As String
    Dim cartellaLogs As String

    ' 06_03 = Formatta data/ora per log
    dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")

    ' 06_04 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()

    ' 06_05 = Costruisci percorso completo file di log
    percorsoLog = cartellaLogs & "\" & nomeFile

    ' 06_06 = Scrive nel file in modalità Append
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        ' 06_07 = Componi e scrivi il testo
        Print #fileNum, dataOra & " | " & nomeFunzione & " | " & messaggio
    Close #fileNum

    ' 06_08 = Uscita normale
    Exit Sub

' 06_09 = Gestione errori interni (silenziosa)
GestioneErrore:
    ' 06_10 = Se fallisce, silenziosamente ignora
    ' Non fare nulla per non interrompere il flusso principale
End Sub

' ********************************************************************************************************************
' *** FINE 06_ATTIVITA = Scrittura log su file personalizzato ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 07_ATTIVITA = Gestione errori interni del logger ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: GestisciErroreInterno
' Descrizione: Gestisce gli errori interni del logger
' Parametri:
'   nomeFunzione      - Nome della funzione che ha generato l'errore
'   NumeroErrore      - Numero dell'errore
'   DescrizioneErrore - Descrizione dell'errore
' Output: Messaggio di errore e scrittura log emergenza
' ====================================================================================================================
Private Sub GestisciErroreInterno(nomeFunzione As String, NumeroErrore As Long, DescrizioneErrore As String)
    ' 07_01 = Mostra messaggio di errore
    MsgBox "ERRORE INTERNO LOGGER:" & vbCrLf & _
           "Funzione: " & nomeFunzione & vbCrLf & _
           "Errore #: " & NumeroErrore & vbCrLf & _
           "Descrizione: " & DescrizioneErrore, _
           vbCritical, "Errore Logger"
    
    ' 07_02 = Tenta scrittura log emergenza
    ScriviLogEmergenza nomeFunzione, NumeroErrore, DescrizioneErrore
End Sub

' ********************************************************************************************************************
' *** FINE 07_ATTIVITA = Gestione errori interni del logger ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 08_ATTIVITA = Scrittura log di emergenza ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: ScriviLogEmergenza
' Descrizione: Scrive un log di emergenza in caso di fallimento del logger
' Parametri:
'   nomeFunzione      - Nome della funzione che ha generato l'errore
'   NumeroErrore      - Numero dell'errore
'   DescrizioneErrore - Descrizione dell'errore
' Output: Scrive una riga di log nel file LogErrore_di_Sistema.txt
' ====================================================================================================================
Private Sub ScriviLogEmergenza(nomeFunzione As String, NumeroErrore As Long, DescrizioneErrore As String)
    ' 08_01 = Gestione errori per la scrittura emergenza
    On Error Resume Next
    
    ' 08_02 = Dichiarazione variabili
    Dim percorsoLog As String
    Dim fileNum As Integer
    Dim dataOra As String
    Dim cartellaLogs As String
    
    ' 08_03 = Formatta data/ora per log
    dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")
    
    ' 08_04 = Imposta percorso cartella (potrebbe non funzionare)
    cartellaLogs = CurrentProject.path & "\Logs"
    
    ' 08_05 = Tenta di creare la cartella (se possibile)
    If Dir(cartellaLogs, vbDirectory) = "" Then
        On Error Resume Next
        MkDir cartellaLogs
        On Error Resume Next
    End If
    
    ' 08_06 = Costruisci percorso file emergenza
    percorsoLog = cartellaLogs & "\LogErrore_di_Sistema.txt"
    
    ' 08_07 = Scrittura in modalità Append
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        ' 08_08 = Scrivi log emergenza
        Print #fileNum, dataOra & " | Scrittura log fallita nella funzione: " & nomeFunzione & _
              " | Errore interno #: " & NumeroErrore & " - " & DescrizioneErrore
    Close #fileNum
    
    ' 08_09 = Se fallisce tutto, tenta scrittura nella root
    On Error Resume Next
    percorsoLog = CurrentProject.path & "\LogErrore_di_Sistema.txt"
    fileNum = FreeFile
    Open percorsoLog For Append As #fileNum
        Print #fileNum, dataOra & " | EMERGENZA: " & nomeFunzione & _
              " | Errore #: " & NumeroErrore & " - " & DescrizioneErrore
    Close #fileNum
End Sub

' ********************************************************************************************************************
' *** FINE 08_ATTIVITA = Scrittura log di emergenza ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 09_ATTIVITA = Utility per pulizia log ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: PulisciLog
' Descrizione: Elimina i file di log più vecchi di N giorni
' Parametri:
'   giorni          - Numero di giorni da mantenere
'   tipoLog         - "ERRORI", "INFO", "TUTTI" (default "TUTTI")
' Output: Elimina i file di log più vecchi
' ====================================================================================================================
Public Sub PulisciLog(Optional giorni As Integer = 30, Optional tipoLog As String = "TUTTI")
    ' 09_01 = Gestione errori
    On Error GoTo GestioneErrore
    
    ' 09_02 = Dichiarazione variabili
    Dim cartellaLogs As String
    Dim nomeFile As String
    Dim dataScadenza As Date
    Dim percorsoFile As String
    Dim fso As Object
    Dim file As Object
    Dim cartella As Object
    
    ' 09_03 = Calcola data scadenza
    dataScadenza = Date - giorni
    
    ' 09_04 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()
    
    ' 09_05 = Crea oggetto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 09_06 = Verifica se la cartella esiste
    If fso.FolderExists(cartellaLogs) Then
        Set cartella = fso.GetFolder(cartellaLogs)
        
        ' 09_07 = Scansiona tutti i file nella cartella
        For Each file In cartella.Files
            ' 09_08 = Verifica se il file è un log
            If InStr(1, file.Name, "Log", vbTextCompare) > 0 Then
                ' 09_09 = Verifica se il file è scaduto
                If file.DateLastModified < dataScadenza Then
                    ' 09_10 = Verifica il tipo di log da eliminare
                    Select Case tipoLog
                        Case "ERRORI"
                            If InStr(1, file.Name, "LogErrori", vbTextCompare) > 0 Then
                                file.Delete
                            End If
                        Case "INFO"
                            If InStr(1, file.Name, "LogInfo", vbTextCompare) > 0 Then
                                file.Delete
                            End If
                        Case Else ' TUTTI
                            file.Delete
                    End Select
                End If
            End If
        Next file
    End If
    
    ' 09_11 = Uscita normale
    Exit Sub

GestioneErrore:
    ' 09_12 = Gestione errori silenziosa
    ' Non fare nulla
End Sub

' ********************************************************************************************************************
' *** FINE 09_ATTIVITA = Utility per pulizia log ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 10_ATTIVITA = Utility per dimensione log ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: DimensioneLog
' Descrizione: Restituisce la dimensione dei file di log
' Parametri:
'   tipoLog         - "ERRORI", "INFO", "TUTTI" (default "TUTTI")
' Output: Dimensione in byte dei file di log
' ====================================================================================================================
Public Function DimensioneLog(Optional tipoLog As String = "TUTTI") As Long
    ' 10_01 = Gestione errori
    On Error GoTo GestioneErrore
    
    ' 10_02 = Dichiarazione variabili
    Dim cartellaLogs As String
    Dim nomeFile As String
    Dim percorsoFile As String
    Dim fso As Object
    Dim file As Object
    Dim cartella As Object
    Dim dimensioneTotale As Long
    
    ' 10_03 = Inizializza dimensione
    dimensioneTotale = 0
    
    ' 10_04 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()
    
    ' 10_05 = Crea oggetto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 10_06 = Verifica se la cartella esiste
    If fso.FolderExists(cartellaLogs) Then
        Set cartella = fso.GetFolder(cartellaLogs)
        
        ' 10_07 = Scansiona tutti i file nella cartella
        For Each file In cartella.Files
            ' 10_08 = Verifica se il file è un log
            If InStr(1, file.Name, "Log", vbTextCompare) > 0 Then
                ' 10_09 = Verifica il tipo di log
                Select Case tipoLog
                    Case "ERRORI"
                        If InStr(1, file.Name, "LogErrori", vbTextCompare) > 0 Then
                            dimensioneTotale = dimensioneTotale + file.size
                        End If
                    Case "INFO"
                        If InStr(1, file.Name, "LogInfo", vbTextCompare) > 0 Then
                            dimensioneTotale = dimensioneTotale + file.size
                        End If
                    Case Else ' TUTTI
                        dimensioneTotale = dimensioneTotale + file.size
                End Select
            End If
        Next file
    End If
    
    ' 10_10 = Restituisci dimensione
    DimensioneLog = dimensioneTotale
    
    ' 10_11 = Uscita normale
    Exit Function

GestioneErrore:
    ' 10_12 = Gestione errori
    DimensioneLog = 0
End Function

' ********************************************************************************************************************
' *** FINE 10_ATTIVITA = Utility per dimensione log ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 11_ATTIVITA = Utility per conteggio log ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: ContaLinieLog
' Descrizione: Conta il numero di righe nei file di log
' Parametri:
'   tipoLog         - "ERRORI", "INFO", "TUTTI" (default "TUTTI")
' Output: Numero di righe nei file di log
' ====================================================================================================================
Public Function ContaLinieLog(Optional tipoLog As String = "TUTTI") As Long
    ' 11_01 = Gestione errori
    On Error GoTo GestioneErrore
    
    ' 11_02 = Dichiarazione variabili
    Dim cartellaLogs As String
    Dim fso As Object
    Dim file As Object
    Dim cartella As Object
    Dim contaLinee As Long
    Dim fileTesto As Object
    
    ' 11_03 = Inizializza contatore
    contaLinee = 0
    
    ' 11_04 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()
    
    ' 11_05 = Crea oggetto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 11_06 = Verifica se la cartella esiste
    If fso.FolderExists(cartellaLogs) Then
        Set cartella = fso.GetFolder(cartellaLogs)
        
        ' 11_07 = Scansiona tutti i file nella cartella
        For Each file In cartella.Files
            ' 11_08 = Verifica se il file è un log
            If InStr(1, file.Name, "Log", vbTextCompare) > 0 Then
                ' 11_09 = Verifica il tipo di log
                Dim contaFile As Boolean
                contaFile = False
                
                Select Case tipoLog
                    Case "ERRORI"
                        If InStr(1, file.Name, "LogErrori", vbTextCompare) > 0 Then
                            contaFile = True
                        End If
                    Case "INFO"
                        If InStr(1, file.Name, "LogInfo", vbTextCompare) > 0 Then
                            contaFile = True
                        End If
                    Case Else ' TUTTI
                        contaFile = True
                End Select
                
                ' 11_10 = Conta le righe del file
                If contaFile Then
                    Set fileTesto = fso.OpenTextFile(file.path, 1) ' 1 = ForReading
                    Do While Not fileTesto.AtEndOfStream
                        fileTesto.ReadLine
                        contaLinee = contaLinee + 1
                    Loop
                    fileTesto.Close
                End If
            End If
        Next file
    End If
    
    ' 11_11 = Restituisci conteggio
    ContaLinieLog = contaLinee
    
    ' 11_12 = Uscita normale
    Exit Function

GestioneErrore:
    ' 11_13 = Gestione errori
    ContaLinieLog = 0
End Function

' ********************************************************************************************************************
' *** FINE 11_ATTIVITA = Utility per conteggio log ***
' ********************************************************************************************************************


' ********************************************************************************************************************
' *** INIZIO 12_ATTIVITA = Utility per visualizzare log ***
' ********************************************************************************************************************

' ====================================================================================================================
' Metodo: VisualizzaLog
' Descrizione: Apre il file di log specificato
' Parametri:
'   tipoLog         - "ERRORI", "INFO" (default "ERRORI")
' Output: Apre il file di log con l'applicazione predefinita
' ====================================================================================================================
Public Sub VisualizzaLog(Optional tipoLog As String = "ERRORI")
    ' 12_01 = Gestione errori
    On Error GoTo GestioneErrore
    
    ' 12_02 = Dichiarazione variabili
    Dim percorsoLog As String
    Dim cartellaLogs As String
    Dim nomeFile As String
    
    ' 12_03 = Ottieni percorso cartella Logs
    cartellaLogs = CreaDirectoryLogs()
    
    ' 12_04 = Determina il file da aprire
    Select Case tipoLog
        Case "ERRORI"
            nomeFile = "LogErrori.txt"
        Case "INFO"
            nomeFile = "LogInfo.txt"
        Case Else
            nomeFile = "LogErrori.txt"
    End Select
    
    ' 12_05 = Costruisci percorso completo
    percorsoLog = cartellaLogs & "\" & nomeFile
    
    ' 12_06 = Verifica se il file esiste
    If Dir(percorsoLog) <> "" Then
        ' 12_07 = Apri con applicazione predefinita
        Application.FollowHyperlink percorsoLog
    Else
        ' 12_08 = File non trovato
        MsgBox "File di log non trovato: " & percorsoLog, vbExclamation, "Log non disponibile"
    End If
    
    ' 12_09 = Uscita normale
    Exit Sub

GestioneErrore:
    ' 12_10 = Gestione errori
    MsgBox "Impossibile aprire il file di log: " & Err.Description, vbCritical, "Errore"
End Sub

' ********************************************************************************************************************
' *** FINE 12_ATTIVITA = Utility per visualizzare log ***
' ********************************************************************************************************************

