VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSys_TB12_Cls01_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'==================================================================================================
'  CLASSE    : MSys_TB12_Cls01_LOG
'  PROGETTO  : MSys_TB12_FILES  (mini-progetto TB12)
'  SCOPO     : logger per il codice applicativo (form, moduli, altre classi): errori con la posizione
'              nel codice, messaggi informativi, file personalizzati, manutenzione dei log.
'              Ogni riga va identica in Debug.Print e nel file.
'  SOSTITUISCE: Msys_TB12_FILES_CLASSE_clsLogger e CLASSE_clsLogger (stessi nomi pubblici)
'  DIPENDENZE: nessuna. La classe si importa da sola in qualsiasi progetto Access.
'  LOG       : <cartella del database>\Logs\<NomeBase>_<TIPO>_AAAA_MM_GG.txt
'              NomeBase predefinito = MSys_TB12_Cls01_LOG ; TIPO = ERRORI | INFO
'              riga = data ora | Origine | Funzione | Pos: SS.PP.ss | Esito | testo
'  VERSIONE  : 1.0 - 2026-11-15
'--------------------------------------------------------------------------------------------------
'  INDICE DELLE SEZIONI, DEI PARAGRAFI E DEI SOTTOPARAGRAFI  (cerca l'etichetta, es. #03.02.01#)
'--------------------------------------------------------------------------------------------------
'  #01#  IMPOSTAZIONI = costanti e variabili
'        #01.01# = costanti          #01.02# = variabili della classe
'  #02#  CICLO_DI_VITA_E_PROPRIETA = inizializzazione e impostazioni
'        #02.01# = Class_Initialize  #02.02# = Class_Terminate
'        #02.03# = proprieta: CartellaLog, NomeBase, Origine, Attivo, ScriviAVideo, UltimoErrore
'  #03#  SCRITTURA = metodi pubblici di scrittura
'        #03.01# = Scrivi                (generico con posizione ed esito)
'        #03.02# = ScriviErrore          #03.02.01# riga ERRORI  #03.02.02# copia nel file INFO
'        #03.03# = ScriviLogErrore       (compatibile: funzione, numero, descrizione, riga)
'        #03.04# = ScriviInfo / ScriviLogEvento (compatibili)
'        #03.05# = ScriviLogPersonalizzato (file scelto dal chiamante)
'        #03.06# = Separatore            (riga di separazione con titolo)
'  #04#  FILE_DI_LOG = percorsi e scrittura fisica
'        #04.01# = CartellaLogs          #04.01.01# cartella impostata o <db>\Logs  #04.01.02# creazione
'        #04.02# = PercorsoFile          (per TIPO e giorno)
'        #04.03# = ScriviRiga            #04.03.01# Debug.Print  #04.03.02# file  #04.03.03# emergenza
'  #05#  MANUTENZIONE = pulizia, dimensione, conteggio, visualizzazione
'        #05.01# = PulisciLog            #05.01.01# scadenza  #05.01.02# scansione ed eliminazione
'        #05.02# = DimensioneLog         #05.03# = ContaRigheLog (alias ContaLinieLog)
'        #05.04# = VisualizzaLog
'  #06#  EMERGENZA = scrittura quando il file normale non e' scrivibile
'        #06.01# = ScriviEmergenza       (%TEMP%\MSys_TB12_Cls01_LOG_EMERGENZA.txt)
'==================================================================================================


'##################################################################################################
'## #01# IMPOSTAZIONI = costanti e variabili
'##################################################################################################

'//#01.01# = costanti
Private Const NOME_CLASSE As String = "MSys_TB12_Cls01_LOG"
Private Const TIPO_ERRORI As String = "ERRORI"
Private Const TIPO_INFO As String = "INFO"

'//#01.02# = variabili della classe
Private m_CartellaLog_s As String
Private m_NomeBase_s As String
Private m_Origine_s As String
Private m_Attivo_b As Boolean
Private m_AVideo_b As Boolean
Private m_UltimoErrore_s As String


'##################################################################################################
'## #02# CICLO_DI_VITA_E_PROPRIETA = inizializzazione e impostazioni
'##################################################################################################

Private Sub Class_Initialize()
    '//#02.01# = valori predefiniti
    m_CartellaLog_s = vbNullString
    m_NomeBase_s = NOME_CLASSE
    m_Origine_s = NOME_CLASSE
    m_Attivo_b = True
    m_AVideo_b = True
End Sub

Private Sub Class_Terminate()
    '//#02.02# = nessuna risorsa aperta: ogni scrittura apre e chiude il file
End Sub

'//#02.03# = proprieta
Public Property Get CartellaLog() As String
    CartellaLog = CartellaLogs()
End Property
Public Property Let CartellaLog(ByVal par_Cartella_s As String)
    m_CartellaLog_s = par_Cartella_s
End Property
Public Property Get NomeBase() As String
    NomeBase = m_NomeBase_s
End Property
Public Property Let NomeBase(ByVal par_Nome_s As String)
    '//nome dei file: <NomeBase>_ERRORI_AAAA_MM_GG.txt (es. il nome del modulo che usa il logger)
    If Len(Trim$(par_Nome_s)) > 0 Then m_NomeBase_s = Trim$(par_Nome_s)
End Property
Public Property Get origine() As String
    origine = m_Origine_s
End Property
Public Property Let origine(ByVal par_Origine_s As String)
    '//testo della colonna "Origine" (es. Form_MASTER); se vuoto resta il nome della classe
    If Len(Trim$(par_Origine_s)) > 0 Then m_Origine_s = Trim$(par_Origine_s)
End Property
Public Property Get Attivo() As Boolean
    Attivo = m_Attivo_b
End Property
Public Property Let Attivo(ByVal par_Valore_b As Boolean)
    m_Attivo_b = par_Valore_b
End Property
Public Property Get ScriviAVideo() As Boolean
    ScriviAVideo = m_AVideo_b
End Property
Public Property Let ScriviAVideo(ByVal par_Valore_b As Boolean)
    m_AVideo_b = par_Valore_b
End Property
Public Property Get UltimoErrore() As String
    UltimoErrore = m_UltimoErrore_s
End Property


'##################################################################################################
'## #03# SCRITTURA = metodi pubblici di scrittura
'##################################################################################################

Public Sub Scrivi(ByVal funzione As String, ByVal posizione As String, ByVal esito As String, ByVal testo As String)
'//#03.01# = riga generica: esito INIZIO, OK, INFO, AVVISO, ERRORE, FINE. ERRORE va nel file ERRORI, il resto in INFO
    If UCase$(esito) = "ERRORE" Then
        ScriviRiga TIPO_ERRORI, funzione, posizione, esito, testo
    Else
        ScriviRiga TIPO_INFO, funzione, posizione, esito, testo
    End If
End Sub

Public Sub ScriviErrore(ByVal funzione As String, ByVal posizione As String, ByVal NumeroErrore As Long, ByVal DescrizioneErrore As String)
'//#03.02# = errore con la posizione nel codice del chiamante (es. "18.06")
Dim sTesto_s As String
    sTesto_s = "errore " & NumeroErrore & ": " & DescrizioneErrore & " -> codice fermo a #" & posizione & "#"
    m_UltimoErrore_s = "#" & posizione & "# " & funzione & " - " & sTesto_s
    '//#03.02.01# = riga nel file ERRORI
    ScriviRiga TIPO_ERRORI, funzione, posizione, "ERRORE", sTesto_s
    '//#03.02.02# = stessa riga nel file INFO, per leggere la sequenza completa in un solo file
    ScriviRiga TIPO_INFO, funzione, posizione, "ERRORE", sTesto_s, False
End Sub

Public Sub ScriviLogErrore(ByVal nomeFunzione As String, ByVal NumeroErrore As Long, ByVal DescrizioneErrore As String, Optional ByVal RigaErrore As Long = 0)
'//#03.03# = compatibile con le versioni precedenti (la riga e' Erl, se il codice ha i numeri di riga)
    ScriviErrore nomeFunzione, IIf(RigaErrore > 0, "riga " & RigaErrore, "-"), NumeroErrore, DescrizioneErrore
End Sub

'//#03.04# = messaggi informativi (compatibili)
Public Sub ScriviInfo(ByVal nomeFunzione As String, ByVal messaggio As String)
    ScriviRiga TIPO_INFO, nomeFunzione, "-", "INFO", messaggio
End Sub
Public Sub ScriviLogEvento(ByVal nomeFunzione As String, ByVal messaggio As String)
    ScriviRiga TIPO_INFO, nomeFunzione, "-", "INFO", messaggio
End Sub

Public Sub ScriviLogPersonalizzato(ByVal nomeFunzione As String, ByVal messaggio As String, ByVal nomeFile As String)
'//#03.05# = riga in un file scelto dal chiamante dentro la cartella Logs (es. "Importazioni.txt")
Dim sRiga_s As String, iF_i As Integer
    If Not m_Attivo_b Then Exit Sub
    On Error GoTo ERRORE
    sRiga_s = ComponiRiga(nomeFunzione, "-", "INFO", messaggio)
    If m_AVideo_b Then Debug.Print sRiga_s
    iF_i = FreeFile
    Open CartellaLogs() & "\" & nomeFile For Append As #iF_i
    Print #iF_i, sRiga_s
    Close #iF_i
    Exit Sub
ERRORE:
    ScriviEmergenza "ScriviLogPersonalizzato", Err.Number, Err.Description & " [" & nomeFile & "]"
End Sub

Public Sub Separatore(Optional ByVal titolo As String = vbNullString)
'//#03.06# = riga di separazione, utile all'inizio di una procedura lunga
    ScriviRiga TIPO_INFO, "----------", "-----", "------", "========== " & titolo & " =========="
End Sub


'##################################################################################################
'## #04# FILE_DI_LOG = percorsi e scrittura fisica
'##################################################################################################

Public Function CartellaLogs() As String
'//#04.01# = cartella dei log
    On Error Resume Next
    '//#04.01.01# = cartella impostata oppure <cartella del database>\Logs
    If Len(m_CartellaLog_s) > 0 Then CartellaLogs = m_CartellaLog_s Else CartellaLogs = CurrentProject.path & "\Logs"
    '//#04.01.02# = creazione se manca
    If Len(Dir$(CartellaLogs, vbDirectory)) = 0 Then MkDir CartellaLogs
End Function

Public Function PercorsoFile(Optional ByVal tipo As String = TIPO_ERRORI, Optional ByVal giorno As Date = 0) As String
'//#04.02# = file del giorno per TIPO (ERRORI o INFO)
    If giorno = 0 Then giorno = Date
    PercorsoFile = CartellaLogs() & "\" & m_NomeBase_s & "_" & UCase$(tipo) & "_" & Format$(giorno, "yyyy_mm_dd") & ".txt"
End Function

Private Function ComponiRiga(ByVal sFunz As String, ByVal sPos As String, ByVal sEsito As String, ByVal sTesto As String) As String
    ComponiRiga = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Origine: " & m_Origine_s & " | Funzione: " & sFunz & _
                  " | Pos: " & Left$(sPos & Space$(8), 8) & " | Esito: " & Left$(sEsito & Space$(6), 6) & " | " & sTesto
End Function

Private Sub ScriviRiga(ByVal sTipo As String, ByVal sFunz As String, ByVal sPos As String, ByVal sEsito As String, ByVal sTesto As String, Optional ByVal bAVideo As Boolean = True)
'//#04.03# = scrittura fisica della riga
Dim sRiga_s As String, iF_i As Integer
    If Not m_Attivo_b Then Exit Sub
    On Error GoTo ERRORE
    sRiga_s = ComponiRiga(sFunz, sPos, sEsito, sTesto)
    '//#04.03.01# = Debug.Print (stessa riga del file)
    If m_AVideo_b And bAVideo Then Debug.Print sRiga_s
    '//#04.03.02# = file del giorno, in coda
    iF_i = FreeFile
    Open PercorsoFile(sTipo) For Append As #iF_i
    Print #iF_i, sRiga_s
    Close #iF_i
    Exit Sub
ERRORE:
    '//#04.03.03# = se il file non e' scrivibile: file di emergenza in %TEMP%
    ScriviEmergenza sFunz, Err.Number, Err.Description & " | riga persa: " & sTesto
End Sub


'##################################################################################################
'## #05# MANUTENZIONE = pulizia, dimensione, conteggio, visualizzazione
'##################################################################################################

Public Function PulisciLog(Optional ByVal giorni As Integer = 30, Optional ByVal filtro As String = "*.txt") As Long
'//#05.01# = elimina i file della cartella Logs piu' vecchi di N giorni; restituisce quanti ne ha eliminati
Dim sCart_s As String, sNome_s As String, dScad As Date, coll As New Collection, v As Variant
    On Error GoTo ERRORE
    '//#05.01.01# = data di scadenza
    dScad = Date - giorni
    sCart_s = CartellaLogs()
    '//#05.01.02# = scansione (prima raccolgo i nomi, poi elimino: Dir non va modificato mentre gira)
    sNome_s = Dir$(sCart_s & "\" & filtro)
    Do While Len(sNome_s) > 0
        If FileDateTime(sCart_s & "\" & sNome_s) < dScad Then coll.Add sCart_s & "\" & sNome_s
        sNome_s = Dir$()
    Loop
    For Each v In coll
        Kill v
        PulisciLog = PulisciLog + 1
    Next v
    ScriviRiga TIPO_INFO, "PulisciLog", "05.01.02", "OK", PulisciLog & " file eliminati (piu' vecchi di " & giorni & " giorni, filtro " & filtro & ")"
    Exit Function
ERRORE:
    ScriviErrore "PulisciLog", "05.01.02", Err.Number, Err.Description
End Function

Public Function DimensioneLog(Optional ByVal filtro As String = "*.txt") As Double
'//#05.02# = byte occupati dai file di log che corrispondono al filtro
Dim sCart_s As String, sNome_s As String
    On Error Resume Next
    sCart_s = CartellaLogs()
    sNome_s = Dir$(sCart_s & "\" & filtro)
    Do While Len(sNome_s) > 0
        DimensioneLog = DimensioneLog + FileLen(sCart_s & "\" & sNome_s)
        sNome_s = Dir$()
    Loop
End Function

Public Function ContaRigheLog(Optional ByVal filtro As String = "*.txt") As Long
'//#05.03# = numero di righe nei file di log che corrispondono al filtro
Dim sCart_s As String, sNome_s As String, coll As New Collection, v As Variant, iF_i As Integer, sRiga_s As String
    On Error Resume Next
    sCart_s = CartellaLogs()
    sNome_s = Dir$(sCart_s & "\" & filtro)
    Do While Len(sNome_s) > 0
        coll.Add sCart_s & "\" & sNome_s
        sNome_s = Dir$()
    Loop
    For Each v In coll
        iF_i = FreeFile
        Open v For Input As #iF_i
        Do While Not EOF(iF_i)
            Line Input #iF_i, sRiga_s
            ContaRigheLog = ContaRigheLog + 1
        Loop
        Close #iF_i
    Next v
End Function
Public Function ContaLinieLog(Optional ByVal filtro As String = "*.txt") As Long
    ContaLinieLog = ContaRigheLog(filtro)
End Function

Public Sub VisualizzaLog(Optional ByVal tipo As String = TIPO_ERRORI)
'//#05.04# = apre il log di oggi del TIPO indicato con l'applicazione predefinita
Dim sFile_s As String
    On Error GoTo ERRORE
    sFile_s = PercorsoFile(tipo)
    If Len(Dir$(sFile_s)) = 0 Then
        MsgBox "Oggi non c'e' ancora un log " & tipo & ":" & vbCrLf & sFile_s, vbInformation, NOME_CLASSE
        Exit Sub
    End If
    Application.FollowHyperlink sFile_s
    Exit Sub
ERRORE:
    ScriviErrore "VisualizzaLog", "05.04", Err.Number, Err.Description
End Sub


'##################################################################################################
'## #06# EMERGENZA = scrittura quando il file normale non e' scrivibile
'##################################################################################################

Private Sub ScriviEmergenza(ByVal sFunz As String, ByVal nErr As Long, ByVal sDescr As String)
'//#06.01# = %TEMP%\MSys_TB12_Cls01_LOG_EMERGENZA.txt ; nessun messaggio a video per non bloccare
Dim iF_i As Integer, sRiga_s As String
    On Error Resume Next
    m_UltimoErrore_s = "#06.01# " & sFunz & " - log non scritto, errore " & nErr & ": " & sDescr
    sRiga_s = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Origine: " & m_Origine_s & " | Funzione: " & sFunz & _
              " | Pos: 06.01    | Esito: ERRORE | log non scritto - errore " & nErr & ": " & sDescr
    Debug.Print sRiga_s
    iF_i = FreeFile
    Open Environ$("TEMP") & "\" & NOME_CLASSE & "_EMERGENZA.txt" For Append As #iF_i
    Print #iF_i, sRiga_s
    Close #iF_i
End Sub
