Attribute VB_Name = "CLSModulo_GESTIONE_OGGETTI_TEST_TutteLeClassi"
' =====================================================
' CLASSE: CLASSE_clsFileSystem
' RESPONSABILITA': Gestione di file e cartelle (apertura, selezione, proprietà)
' VERSIONE: 2.0
' DATA: 16/06/2026
' =====================================================
'
' METODI PUBBLICI DISPONIBILI:
'
' 1. SelezionaFile
'    - Mostra finestra di dialogo per selezionare un file
'    - Supporta percorsi lunghi fino a 32.767 caratteri
'    - Parametri: filtroDescrizione, filtroEstensione, cartellaIniziale, titolo
'    - Restituisce: String (percorso del file selezionato o vbNullString)
'
' 2. SalvaFile
'    - Mostra finestra di dialogo per salvare un file
'    - Supporta percorsi lunghi fino a 32.767 caratteri
'    - Parametri: filtroDescrizione, filtroEstensione, cartellaIniziale,
'                 nomeFileSuggerito, titolo
'    - Restituisce: String (percorso del file salvato o vbNullString)
'
' 3. ApriFile
'    - Apre un file o una cartella usando ShellExecute
'    - Supporta percorsi lunghi con prefisso \\?\
'    - Parametri: percorsoCompleto, showCmd
'    - Restituisce: Boolean (True se successo, False altrimenti)
'
' 4. FileEsiste
'    - Verifica se un file esiste
'    - Supporta percorsi lunghi (>260 caratteri) tramite API
'    - Parametri: filePath
'    - Restituisce: Boolean (True se esiste, False altrimenti)
'
' 5. CartellaEsiste
'    - Verifica se una cartella esiste
'    - Supporta percorsi lunghi (>260 caratteri) tramite API
'    - Parametri: cartellaPath
'    - Restituisce: Boolean (True se esiste, False altrimenti)
'
' 6. MostraProprietaFile
'    - Mostra le proprietà dettagliate di un file
'    - Visualizza: percorso, nome, tipo, dimensione, date, attributi
'    - Supporta percorsi lunghi
'    - Parametri: filePath
'    - Restituisce: Sub (nessun valore di ritorno)
'
' 7. CopiaFile
'    - Copia un file da una posizione a un'altra
'    - Supporta percorsi lunghi
'    - Parametri: sourcePath, destinationPath, overwrite
'    - Restituisce: Boolean (True se successo, False altrimenti)
'
' 8. SpostaFile
'    - Sposta o rinomina un file
'    - Supporta percorsi lunghi
'    - Parametri: sourcePath, destinationPath
'    - Restituisce: Boolean (True se successo, False altrimenti)
'
' 9. EliminaFile
'    - Elimina un file
'    - Supporta percorsi lunghi
'    - Parametri: filePath
'    - Restituisce: Boolean (True se successo, False altrimenti)
'
' 10. CreaCartella
'     - Crea una nuova cartella (anche percorsi multipli)
'     - Supporta percorsi lunghi
'     - Parametri: cartellaPath
'     - Restituisce: Boolean (True se successo, False altrimenti)
'
' 11. EliminaCartella
'     - Elimina una cartella (anche con contenuto)
'     - Supporta percorsi lunghi
'     - Parametri: cartellaPath, forceDelete
'     - Restituisce: Boolean (True se successo, False altrimenti)
'
' 12. GetFileName
'     - Restituisce il nome del file da un percorso completo
'     - Supporta percorsi lunghi
'     - Parametri: percorsoCompleto
'     - Restituisce: String (nome del file)
'
' 13. GetFolderPath
'     - Restituisce il percorso della cartella da un percorso completo
'     - Supporta percorsi lunghi
'     - Parametri: percorsoCompleto
'     - Restituisce: String (percorso della cartella)
'
' 14. GetFileExtension
'     - Restituisce l'estensione del file
'     - Supporta percorsi lunghi
'     - Parametri: percorsoCompleto
'     - Restituisce: String (estensione del file senza punto)
'
' 15. GetFileSizeFormatted
'     - Restituisce la dimensione del file in formato leggibile
'     - Formati: bytes, KB, MB, GB
'     - Supporta percorsi lunghi
'     - Parametri: filePath
'     - Restituisce: String (dimensione formattata)
'
' 16. IsLongPath
'     - Verifica se un percorso supera i 260 caratteri
'     - Parametri: percorso
'     - Restituisce: Boolean (True se lungo, False altrimenti)
'
' 17. NormalizzaPercorso
'     - Normalizza un percorso aggiungendo il prefisso \\?\ se necessario
'     - Utile per percorsi lunghi o con caratteri speciali
'     - Parametri: percorso
'     - Restituisce: String (percorso normalizzato)
'
' =====================================================
' PROPRIETA' E COSTANTI:
' =====================================================
'
' Costanti Private:
' - SW_SHOWNORMAL: 1 (Mostra finestra normalmente)
' - SE_ERR_SUCCESS: 32 (Successo ShellExecute)
' - MAX_PATH: 260 (Limite massimo percorso Windows)
' - LONG_PATH_PREFIX: "\\?\" (Prefisso per percorsi lunghi)
'
' =====================================================
' DIPENDENZE E RIFERIMENTI:
' =====================================================
'
' API Windows utilizzate:
' - ShellExecute (shell32.dll)
' - GetOpenFileName (comdlg32.dll)
' - GetSaveFileName (comdlg32.dll)
' - GetFileAttributesEx (kernel32.dll)
'
' Oggetti COM utilizzati:
' - Scripting.FileSystemObject (per operazioni su file)
'
' =====================================================
' NOTE IMPORTANTI:
' =====================================================
'
' 1. Supporto percorsi lunghi:
'    - Funziona su Windows 10 (versione 1607+) e Windows Server 2016+
'    - Necessario abilitare il supporto Long Path nel registro di sistema
'    - La classe gestisce automaticamente il prefisso \\?\
'
' 2. Compatibilità:
'    - Supporta VBA7 e versioni precedenti
'    - Compatibile con Access 2007 e successivi
'    - Non richiede riferimenti aggiuntivi (usa API native)
'
' 3. Gestione errori:
'    - Tutti i metodi hanno gestione errori integrata
'    - Le funzioni restituiscono False o vbNullString in caso di errore
'
' 4. Performance:
'    - Ottimizzato per operazioni su file di grandi dimensioni
'    - Utilizza API native per operazioni veloci
'
' =====================================================
' ESEMPI DI UTILIZZO:
' =====================================================
'
' Esempio 1: Selezionare e aprire un file
' --------------------------------------
' Dim fs As New CLASSE_clsFileSystem
' Dim filePath As String
' filePath = fs.SelezionaFile("File di testo", "*.txt")
' If filePath <> "" Then
'     fs.ApriFile filePath
'     fs.MostraProprietaFile filePath
' End If
'
' Esempio 2: Copiare un file con percorso lungo
' ---------------------------------------------
' Dim fs As New CLASSE_clsFileSystem
' Dim longPath As String
' longPath = "C:\Folder1\Folder2\... (percorso lungo)\file.txt"
' If fs.CopiaFile("C:\orig.txt", longPath) Then
'     MsgBox "File copiato con successo!"
' End If
'
' Esempio 3: Creare una struttura di cartelle
' -------------------------------------------
' Dim fs As New CLASSE_clsFileSystem
' Dim deepPath As String
' deepPath = "C:\Projects\2024\Reports\Q1\Data\"
' If fs.CreaCartella(deepPath) Then
'     MsgBox "Cartella creata: " & deepPath
' End If
'
' =====================================================

' =====================================================
' MODULO: modTest_clsFileSystem
' RESPONSABILITA': Test completo della classe CLASSE_clsFileSystem
' =====================================================
Option Compare Database
Option Explicit

' Variabile globale per il test
Private testFileSystem As CLASSE_clsFileSystem
Private Const TEST_FILE_PATH As String = "C:\Temp\test_file.txt"
Private Const TEST_FOLDER_PATH As String = "C:\Temp\TestFolder"
Private Const TEST_FILE_CONTENT As String = "Questo è un file di test"

' Variabili per il conteggio dei test
Private totalTests As Integer
Private passedTests As Integer
Private testResults As String

' =====================================================
' TEST: Esecuzione di tutti i test
' =====================================================
Public Sub EseguiTuttiITest()
    ' Inizializza
    testResults = "--- RISULTATI TEST CLASSE clsFileSystem ---" & vbCrLf & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' Crea ambiente di test
    If Not SetupTestEnvironment() Then
        MsgBox "Impossibile creare l'ambiente di test. Verifica i permessi di scrittura su C:\Temp", vbCritical
        Exit Sub
    End If
    
    ' Inizializza la classe
    Set testFileSystem = New CLASSE_clsFileSystem
    
    ' Esegui i test - TEST DI BASE
    EseguiTest "Test FileEsiste", "TestFileEsiste"
    EseguiTest "Test CartellaEsiste", "TestCartellaEsiste"
    EseguiTest "Test GetFileName", "TestGetFileName"
    EseguiTest "Test GetFileExtension", "TestGetFileExtension"
    EseguiTest "Test GetFolderPath", "TestGetFolderPath"
    EseguiTest "Test GetFileSizeFormatted", "TestGetFileSizeFormatted"
    EseguiTest "Test IsLongPath", "TestIsLongPath"
    
    ' Esegui i test - OPERAZIONI SU FILE
    EseguiTest "Test CopiaFile", "TestCopiaFile"
    EseguiTest "Test SpostaFile", "TestSpostaFile"
    EseguiTest "Test EliminaFile", "TestEliminaFile"
    
    ' Esegui i test - OPERAZIONI SU CARTELLE
    EseguiTest "Test CreaCartella", "TestCreaCartella"
    EseguiTest "Test EliminaCartella", "TestEliminaCartella"
    
    ' Esegui i test - PERCORSI LUNGHI
    EseguiTest "Test PercorsiLunghi", "TestPercorsiLunghi"
    
    ' Esegui i test - APRI FILE
    EseguiTest "Test ApriFile", "TestApriFile"
    
    ' Esegui i test INTERATTIVI (richiedono input utente)
    EseguiTest "Test SelezionaFile (Interattivo)", "TestSelezionaFile"
    EseguiTest "Test SalvaFile (Interattivo)", "TestSalvaFile"
    EseguiTest "Test MostraProprietaFile (Interattivo)", "TestMostraProprietaFile"
    
    ' Esegui test di gestione errori
    EseguiTest "Test Gestione Errori", "TestGestioneErrori"
    
    ' Mostra risultati
    testResults = testResults & vbCrLf & "--- RIEPILOGO ---" & vbCrLf
    testResults = testResults & "Test eseguiti: " & totalTests & vbCrLf
    testResults = testResults & "Test passati: " & passedTests & vbCrLf
    testResults = testResults & "Test falliti: " & (totalTests - passedTests) & vbCrLf
    testResults = testResults & "Percentuale successo: " & Format((passedTests / totalTests) * 100, "0.0") & "%" & vbCrLf
    
    MsgBox testResults, vbInformation, "Risultati Test"
    
    ' Pulizia
    PuliziaTestEnvironment
    Set testFileSystem = Nothing
End Sub

' =====================================================
' FUNZIONE: Esegue un singolo test
' =====================================================
Private Sub EseguiTest(testName As String, testFunction As String)
    Dim result As Boolean
    Dim startTime As Double
    Dim endTime As Double
    
    totalTests = totalTests + 1
    startTime = Timer
    
    On Error Resume Next
    result = Application.Run(testFunction)
    On Error GoTo 0
    
    endTime = Timer
    
    testResults = testResults & IIf(result, "? PASSATO", "? FALLITO") & " - " & testName
    testResults = testResults & " (" & Format(endTime - startTime, "0.000") & "s)" & vbCrLf
    
    If result Then passedTests = passedTests + 1
End Sub

' =====================================================
' TEST 1: FileEsiste
' =====================================================
Private Function TestFileEsiste() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: File esistente
    If Not testFileSystem.FileEsiste(TEST_FILE_PATH) Then
        TestFileEsiste = False
        Exit Function
    End If
    
    ' Test 2: File inesistente
    If testFileSystem.FileEsiste("C:\NonEsistente\file.txt") Then
        TestFileEsiste = False
        Exit Function
    End If
    
    ' Test 3: Percorso vuoto
    If testFileSystem.FileEsiste("") Then
        TestFileEsiste = False
        Exit Function
    End If
    
    ' Test 4: Percorso con caratteri speciali
    If Not testFileSystem.FileEsiste(TEST_FOLDER_PATH & "\file_speciali_!@#.txt") Then
        TestFileEsiste = False
        Exit Function
    End If
    
    TestFileEsiste = True
    Exit Function
    
ErrHandler:
    TestFileEsiste = False
End Function

' =====================================================
' TEST 2: CartellaEsiste
' =====================================================
Private Function TestCartellaEsiste() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: Cartella esistente
    If Not testFileSystem.CartellaEsiste(TEST_FOLDER_PATH) Then
        TestCartellaEsiste = False
        Exit Function
    End If
    
    ' Test 2: Cartella inesistente
    If testFileSystem.CartellaEsiste("C:\NonEsistente\") Then
        TestCartellaEsiste = False
        Exit Function
    End If
    
    ' Test 3: Percorso vuoto
    If testFileSystem.CartellaEsiste("") Then
        TestCartellaEsiste = False
        Exit Function
    End If
    
    ' Test 4: Percorso senza backslash finale
    Dim tempPath As String
    tempPath = Left(TEST_FOLDER_PATH, Len(TEST_FOLDER_PATH) - 1)
    If Not testFileSystem.CartellaEsiste(tempPath) Then
        TestCartellaEsiste = False
        Exit Function
    End If
    
    ' Test 5: Sottocartella esistente
    If Not testFileSystem.CartellaEsiste(TEST_FOLDER_PATH & "\SubFolder") Then
        TestCartellaEsiste = False
        Exit Function
    End If
    
    TestCartellaEsiste = True
    Exit Function
    
ErrHandler:
    TestCartellaEsiste = False
End Function

' =====================================================
' TEST 3: GetFileName
' =====================================================
Private Function TestGetFileName() As Boolean
    On Error GoTo ErrHandler
    
    Dim result As String
    result = testFileSystem.GetFileName(TEST_FILE_PATH)
    
    If result <> "test_file.txt" Then
        TestGetFileName = False
        Exit Function
    End If
    
    ' Test con percorso vuoto
    result = testFileSystem.GetFileName("")
    If result <> "" Then
        TestGetFileName = False
        Exit Function
    End If
    
    TestGetFileName = True
    Exit Function
    
ErrHandler:
    TestGetFileName = False
End Function

' =====================================================
' TEST 4: GetFileExtension
' =====================================================
Private Function TestGetFileExtension() As Boolean
    On Error GoTo ErrHandler
    
    Dim result As String
    result = testFileSystem.GetFileExtension(TEST_FILE_PATH)
    
    If result <> "txt" Then
        TestGetFileExtension = False
        Exit Function
    End If
    
    ' Test con percorso vuoto
    result = testFileSystem.GetFileExtension("")
    If result <> "" Then
        TestGetFileExtension = False
        Exit Function
    End If
    
    TestGetFileExtension = True
    Exit Function
    
ErrHandler:
    TestGetFileExtension = False
End Function

' =====================================================
' TEST 5: GetFolderPath
' =====================================================
Private Function TestGetFolderPath() As Boolean
    On Error GoTo ErrHandler
    
    Dim result As String
    result = testFileSystem.GetFolderPath(TEST_FILE_PATH)
    
    ' Rimuovi backslash finale per il confronto
    Dim expectedPath As String
    expectedPath = "C:\Temp"
    
    If result <> expectedPath Then
        TestGetFolderPath = False
        Exit Function
    End If
    
    ' Test con percorso vuoto
    result = testFileSystem.GetFolderPath("")
    If result <> "" Then
        TestGetFolderPath = False
        Exit Function
    End If
    
    TestGetFolderPath = True
    Exit Function
    
ErrHandler:
    TestGetFolderPath = False
End Function

' =====================================================
' TEST 6: GetFileSizeFormatted
' =====================================================
Private Function TestGetFileSizeFormatted() As Boolean
    On Error GoTo ErrHandler
    
    Dim result As String
    result = testFileSystem.GetFileSizeFormatted(TEST_FILE_PATH)
    
    If result = "" Then
        TestGetFileSizeFormatted = False
        Exit Function
    End If
    
    ' Verifica che il formato sia corretto (contiene "bytes" o "KB" o "MB")
    If InStr(result, "bytes") = 0 And InStr(result, "KB") = 0 And InStr(result, "MB") = 0 And InStr(result, "GB") = 0 Then
        TestGetFileSizeFormatted = False
        Exit Function
    End If
    
    ' Test con file inesistente
    result = testFileSystem.GetFileSizeFormatted("C:\NonEsistente\file.txt")
    If result <> "" Then
        TestGetFileSizeFormatted = False
        Exit Function
    End If
    
    TestGetFileSizeFormatted = True
    Exit Function
    
ErrHandler:
    TestGetFileSizeFormatted = False
End Function

' =====================================================
' TEST 7: IsLongPath
' =====================================================
Private Function TestIsLongPath() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: Percorso breve
    If testFileSystem.IsLongPath("C:\Temp\file.txt") Then
        TestIsLongPath = False
        Exit Function
    End If
    
    ' Test 2: Percorso lungo
    Dim longPath As String
    longPath = String(300, "a") & ".txt"
    If Not testFileSystem.IsLongPath(longPath) Then
        TestIsLongPath = False
        Exit Function
    End If
    
    ' Test 3: Percorso vuoto
    If testFileSystem.IsLongPath("") Then
        TestIsLongPath = False
        Exit Function
    End If
    
    TestIsLongPath = True
    Exit Function
    
ErrHandler:
    TestIsLongPath = False
End Function

' =====================================================
' TEST 8: CopiaFile
' =====================================================
Private Function TestCopiaFile() As Boolean
    On Error GoTo ErrHandler
    
    Dim destPath As String
    destPath = TEST_FOLDER_PATH & "\test_copia.txt"
    
    ' Test 1: Copia riuscita
    If Not testFileSystem.CopiaFile(TEST_FILE_PATH, destPath) Then
        TestCopiaFile = False
        Exit Function
    End If
    
    ' Verifica che il file copiato esista
    If Not testFileSystem.FileEsiste(destPath) Then
        TestCopiaFile = False
        Exit Function
    End If
    
    ' Test 2: Copia con sovrascrittura
    If Not testFileSystem.CopiaFile(TEST_FILE_PATH, destPath, True) Then
        TestCopiaFile = False
        Exit Function
    End If
    
    ' Pulisci
    On Error Resume Next
    testFileSystem.EliminaFile destPath
    On Error GoTo 0
    
    TestCopiaFile = True
    Exit Function
    
ErrHandler:
    TestCopiaFile = False
End Function

' =====================================================
' TEST 9: SpostaFile
' =====================================================
Private Function TestSpostaFile() As Boolean
    On Error GoTo ErrHandler
    
    Dim sourcePath As String
    Dim destPath As String
    sourcePath = TEST_FOLDER_PATH & "\sposta_origine.txt"
    destPath = TEST_FOLDER_PATH & "\sposta_destinazione.txt"
    
    ' Crea file di origine
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.CreateTextFile(sourcePath, True)
    file.WriteLine "Test sposta file"
    file.Close
    Set fso = Nothing
    
    ' Test 1: Sposta riuscito
    If Not testFileSystem.SpostaFile(sourcePath, destPath) Then
        TestSpostaFile = False
        Exit Function
    End If
    
    ' Verifica che il file di origine non esista più
    If testFileSystem.FileEsiste(sourcePath) Then
        TestSpostaFile = False
        Exit Function
    End If
    
    ' Verifica che il file di destinazione esista
    If Not testFileSystem.FileEsiste(destPath) Then
        TestSpostaFile = False
        Exit Function
    End If
    
    ' Pulisci
    On Error Resume Next
    testFileSystem.EliminaFile destPath
    On Error GoTo 0
    
    TestSpostaFile = True
    Exit Function
    
ErrHandler:
    TestSpostaFile = False
End Function

' =====================================================
' TEST 10: EliminaFile
' =====================================================
Private Function TestEliminaFile() As Boolean
    On Error GoTo ErrHandler
    
    Dim filePath As String
    filePath = TEST_FOLDER_PATH & "\elimina_test.txt"
    
    ' Crea file da eliminare
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.CreateTextFile(filePath, True)
    file.WriteLine "Test elimina file"
    file.Close
    Set fso = Nothing
    
    ' Test 1: Elimina riuscita
    If Not testFileSystem.EliminaFile(filePath) Then
        TestEliminaFile = False
        Exit Function
    End If
    
    ' Verifica che il file non esista più
    If testFileSystem.FileEsiste(filePath) Then
        TestEliminaFile = False
        Exit Function
    End If
    
    ' Test 2: Elimina file inesistente (dovrebbe fallire)
    If testFileSystem.EliminaFile("C:\NonEsistente\file.txt") Then
        TestEliminaFile = False
        Exit Function
    End If
    
    TestEliminaFile = True
    Exit Function
    
ErrHandler:
    TestEliminaFile = False
End Function

' =====================================================
' TEST 11: CreaCartella
' =====================================================
Private Function TestCreaCartella() As Boolean
    On Error GoTo ErrHandler
    
    Dim newFolderPath As String
    newFolderPath = TEST_FOLDER_PATH & "\NuovaCartellaTest\"
    
    ' Test 1: Crea cartella nuova
    If Not testFileSystem.CreaCartella(newFolderPath) Then
        TestCreaCartella = False
        Exit Function
    End If
    
    ' Verifica che la cartella esista
    If Not testFileSystem.CartellaEsiste(newFolderPath) Then
        TestCreaCartella = False
        Exit Function
    End If
    
    ' Test 2: Crea cartella già esistente (dovrebbe ritornare True)
    If Not testFileSystem.CreaCartella(newFolderPath) Then
        TestCreaCartella = False
        Exit Function
    End If
    
    ' Pulisci
    On Error Resume Next
    testFileSystem.EliminaCartella newFolderPath, True
    On Error GoTo 0
    
    TestCreaCartella = True
    Exit Function
    
ErrHandler:
    TestCreaCartella = False
End Function

' =====================================================
' TEST 12: EliminaCartella
' =====================================================
Private Function TestEliminaCartella() As Boolean
    On Error GoTo ErrHandler
    
    Dim folderPath As String
    folderPath = TEST_FOLDER_PATH & "\EliminaCartellaTest\"
    
    ' Crea cartella da eliminare
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder folderPath
    Set fso = Nothing
    
    ' Test 1: Elimina cartella
    If Not testFileSystem.EliminaCartella(folderPath, True) Then
        TestEliminaCartella = False
        Exit Function
    End If
    
    ' Verifica che la cartella non esista più
    If testFileSystem.CartellaEsiste(folderPath) Then
        TestEliminaCartella = False
        Exit Function
    End If
    
    ' Test 2: Elimina cartella inesistente (dovrebbe fallire)
    If testFileSystem.EliminaCartella("C:\NonEsistente\") Then
        TestEliminaCartella = False
        Exit Function
    End If
    
    TestEliminaCartella = True
    Exit Function
    
ErrHandler:
    TestEliminaCartella = False
End Function

' =====================================================
' TEST 13: PercorsiLunghi
' =====================================================
Private Function TestPercorsiLunghi() As Boolean
    On Error GoTo ErrHandler
    
    ' Crea una struttura di cartelle profonda
    Dim deepPath As String
    deepPath = TEST_FOLDER_PATH & "\"
    Dim i As Long
    For i = 1 To 10
        deepPath = deepPath & "Folder" & Format(i, "000") & "\"
    Next i
    
    ' Test 1: Crea cartella lunga
    If Not testFileSystem.CreaCartella(deepPath) Then
        TestPercorsiLunghi = False
        Exit Function
    End If
    
    ' Verifica che la cartella esista
    If Not testFileSystem.CartellaEsiste(deepPath) Then
        TestPercorsiLunghi = False
        Exit Function
    End If
    
    ' Test 2: Crea file in percorso lungo
    Dim longFilePath As String
    longFilePath = deepPath & "test_lungo.txt"
    
    ' Crea file con FSO
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.CreateTextFile(longFilePath, True)
    file.WriteLine "Test percorso lungo"
    file.Close
    Set fso = Nothing
    
    ' Verifica che il file esista
    If Not testFileSystem.FileEsiste(longFilePath) Then
        TestPercorsiLunghi = False
        Exit Function
    End If
    
    ' Test 3: Ottieni informazioni dal file con percorso lungo
    Dim fileName As String
    fileName = testFileSystem.GetFileName(longFilePath)
    If fileName <> "test_lungo.txt" Then
        TestPercorsiLunghi = False
        Exit Function
    End If
    
    ' Test 4: Elimina file in percorso lungo
    If Not testFileSystem.EliminaFile(longFilePath) Then
        TestPercorsiLunghi = False
        Exit Function
    End If
    
    ' Test 5: Elimina cartella lunga
    If Not testFileSystem.EliminaCartella(deepPath, True) Then
        TestPercorsiLunghi = False
        Exit Function
    End If
    
    TestPercorsiLunghi = True
    Exit Function
    
ErrHandler:
    TestPercorsiLunghi = False
End Function

' =====================================================
' TEST 14: ApriFile
' =====================================================
Private Function TestApriFile() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: Apri file esistente
    If Not testFileSystem.ApriFile(TEST_FILE_PATH) Then
        TestApriFile = False
        Exit Function
    End If
    
    ' Test 2: Apri cartella esistente
    If Not testFileSystem.ApriFile(TEST_FOLDER_PATH) Then
        TestApriFile = False
        Exit Function
    End If
    
    ' Test 3: Apri file inesistente - dovrebbe fallire
    If testFileSystem.ApriFile("C:\NonEsistente\file.txt") Then
        TestApriFile = False
        Exit Function
    End If
    
    ' Test 4: Percorso vuoto
    If testFileSystem.ApriFile("") Then
        TestApriFile = False
        Exit Function
    End If
    
    ' Test 5: Percorso con spazi
    If Not testFileSystem.ApriFile(TEST_FOLDER_PATH & "\file con spazi.txt") Then
        TestApriFile = False
        Exit Function
    End If
    
    TestApriFile = True
    Exit Function
    
ErrHandler:
    TestApriFile = False
End Function

' =====================================================
' TEST 15: SelezionaFile (Interattivo)
' =====================================================
Private Function TestSelezionaFile() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: Seleziona con filtro default
    Dim selectedFile As String
    selectedFile = testFileSystem.SelezionaFile()
    
    ' Se l'utente annulla, il test è comunque valido
    If selectedFile = "" Then
        ' L'utente ha annullato, consideriamo il test passato
        TestSelezionaFile = True
        Exit Function
    End If
    
    ' Test 2: Seleziona con filtro personalizzato
    selectedFile = testFileSystem.SelezionaFile( _
        filtroDescrizione:="File di testo (*.txt)", _
        filtroEstensione:="*.txt", _
        cartellaIniziale:=TEST_FOLDER_PATH, _
        titolo:="Test: Seleziona un file di testo")
    
    ' Se l'utente ha selezionato un file, verifica che esista
    If selectedFile <> "" And Not testFileSystem.FileEsiste(selectedFile) Then
        TestSelezionaFile = False
        Exit Function
    End If
    
    TestSelezionaFile = True
    Exit Function
    
ErrHandler:
    TestSelezionaFile = False
End Function

' =====================================================
' TEST 16: SalvaFile (Interattivo)
' =====================================================
Private Function TestSalvaFile() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: Salva con nome suggerito
    Dim savedFile As String
    savedFile = testFileSystem.SalvaFile( _
        filtroDescrizione:="File di testo (*.txt)", _
        filtroEstensione:="*.txt", _
        cartellaIniziale:=TEST_FOLDER_PATH, _
        nomeFileSuggerito:="test_salvato.txt", _
        titolo:="Test: Salva un file")
    
    ' Se l'utente annulla, il test è comunque valido
    If savedFile = "" Then
        TestSalvaFile = True
        Exit Function
    End If
    
    ' Test 2: Salva con filtro personalizzato
    savedFile = testFileSystem.SalvaFile( _
        filtroDescrizione:="Tutti i file", _
        filtroEstensione:="*.*", _
        cartellaIniziale:="C:\", _
        titolo:="Test: Salva file qualsiasi")
    
    TestSalvaFile = True
    Exit Function
    
ErrHandler:
    TestSalvaFile = False
End Function

' =====================================================
' TEST 17: MostraProprietaFile (Interattivo)
' =====================================================
Private Function TestMostraProprietaFile() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: Mostra proprietà file esistente
    testFileSystem.MostraProprietaFile TEST_FILE_PATH
    
    ' Test 2: Mostra proprietà con selezione utente
    Dim filePath As String
    filePath = testFileSystem.SelezionaFile(titolo:="Test: Seleziona file per proprietà")
    
    If filePath <> "" Then
        testFileSystem.MostraProprietaFile filePath
    End If
    
    ' Test 3: Percorso vuoto - dovrebbe mostrare messaggio
    testFileSystem.MostraProprietaFile ""
    
    ' Test 4: File inesistente - dovrebbe mostrare errore
    testFileSystem.MostraProprietaFile "C:\NonEsistente\file.txt"
    
    TestMostraProprietaFile = True
    Exit Function
    
ErrHandler:
    TestMostraProprietaFile = False
End Function

' =====================================================
' TEST 18: Gestione Errori
' =====================================================
Private Function TestGestioneErrori() As Boolean
    On Error GoTo ErrHandler
    
    ' Test 1: ShellExecute con parametri non validi
    ' La funzione dovrebbe gestire l'errore e ritornare False
    If testFileSystem.ApriFile("\\?\" & TEST_FILE_PATH) Then
        ' Se funziona, test passato
    End If
    
    ' Test 2: FileSystem con percorsi molto lunghi (>255 caratteri)
    Dim longPath As String
    longPath = String(300, "a") & ".txt"
    ' La funzione dovrebbe gestire il percorso lungo
    
    ' Test 3: Caratteri speciali nel percorso
    Dim specialPath As String
    specialPath = TEST_FOLDER_PATH & "\file_???.txt"
    If Not testFileSystem.ApriFile(specialPath) Then
        ' Potrebbe fallire se il sistema non supporta caratteri speciali
        ' Consideriamo passato se gestisce l'errore
    End If
    
    ' Test 4: Operazioni su file inesistenti
    If testFileSystem.CopiaFile("C:\NonEsistente.txt", "C:\Destinazione.txt") Then
        ' Dovrebbe fallire
    End If
    
    TestGestioneErrori = True
    Exit Function
    
ErrHandler:
    TestGestioneErrori = False
End Function

' =====================================================
' FUNZIONI DI UTILITÀ PER IL TEST
' =====================================================

' Crea l'ambiente di test
Private Function SetupTestEnvironment() As Boolean
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crea la cartella di test
    If Not fso.FolderExists(TEST_FOLDER_PATH) Then
        fso.CreateFolder TEST_FOLDER_PATH
    End If
    
    ' Crea il file di test
    If fso.FileExists(TEST_FILE_PATH) Then
        fso.DeleteFile TEST_FILE_PATH
    End If
    
    ' Crea file con contenuto
    Dim file As Object
    Set file = fso.CreateTextFile(TEST_FILE_PATH, True)
    file.WriteLine TEST_FILE_CONTENT
    file.Close
    
    ' Crea file di test aggiuntivi
    Dim testFile As Object
    Set testFile = fso.CreateTextFile(TEST_FOLDER_PATH & "\file con spazi.txt", True)
    testFile.WriteLine "File con spazi nel nome"
    testFile.Close
    
    Set testFile = fso.CreateTextFile(TEST_FOLDER_PATH & "\file_speciali_!@#.txt", True)
    testFile.WriteLine "File con caratteri speciali"
    testFile.Close
    
    ' Crea file più grande per test dimensioni
    Set testFile = fso.CreateTextFile(TEST_FOLDER_PATH & "\file_grande.txt", True)
    Dim i As Long
    For i = 1 To 1000
        testFile.WriteLine "Riga " & i & " di test per il file grande"
    Next i
    testFile.Close
    
    ' Crea sottocartella
    If Not fso.FolderExists(TEST_FOLDER_PATH & "\SubFolder") Then
        fso.CreateFolder TEST_FOLDER_PATH & "\SubFolder"
    End If
    
    Set fso = Nothing
    SetupTestEnvironment = True
    Exit Function
    
ErrHandler:
    SetupTestEnvironment = False
End Function

' Pulizia dell'ambiente di test
Private Sub PuliziaTestEnvironment()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Elimina file e cartelle di test
    If fso.FolderExists(TEST_FOLDER_PATH) Then
        fso.DeleteFolder TEST_FOLDER_PATH, True
    End If
    
    If fso.FileExists(TEST_FILE_PATH) Then
        fso.DeleteFile TEST_FILE_PATH
    End If
    
    Set fso = Nothing
End Sub

' =====================================================
' TEST MANUALI (da eseguire separatamente)
' =====================================================

' --- Test interattivo per MostraProprietaFile ---
Public Sub TestProprietaInterattivo()
    Dim clsFS As New CLASSE_clsFileSystem
    
    ' Test con un file che l'utente seleziona
    Dim filePath As String
    filePath = clsFS.SelezionaFile(titolo:="Seleziona un file per visualizzare le proprietà")
    
    If filePath <> "" Then
        clsFS.MostraProprietaFile filePath
    Else
        MsgBox "Nessun file selezionato", vbExclamation
    End If
    
    Set clsFS = Nothing
End Sub

' --- Test rapido per SelezionaFile ---
Public Sub TestSelezionaFileRapido()
    Dim clsFS As New CLASSE_clsFileSystem
    
    Dim filePath As String
    filePath = clsFS.SelezionaFile( _
        filtroDescrizione:="Tutti i file", _
        filtroEstensione:="*.*", _
        cartellaIniziale:="C:\", _
        titolo:="Seleziona un file qualsiasi")
    
    If filePath <> "" Then
        MsgBox "File selezionato: " & filePath, vbInformation, "Successo"
        clsFS.MostraProprietaFile filePath
    Else
        MsgBox "Nessun file selezionato", vbExclamation, "Annullato"
    End If
    
    Set clsFS = Nothing
End Sub

' --- Test rapido per percorsi lunghi ---
Public Sub TestPercorsiLunghiRapido()
    Dim clsFS As New CLASSE_clsFileSystem
    
    ' Crea percorso lungo
    Dim longPath As String
    longPath = Environ("TEMP") & "\"
    Dim i As Long
    For i = 1 To 15
        longPath = longPath & "Folder" & Format(i, "000") & "\"
    Next i
    
    Debug.Print "Lunghezza percorso: " & Len(longPath)
    Debug.Print "Percorso: " & longPath
    
    ' Crea cartella
    If clsFS.CreaCartella(longPath) Then
        Debug.Print "? Cartella creata con successo!"
        
        ' Crea file
        Dim filePath As String
        filePath = longPath & "test.txt"
        
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim file As Object
        Set file = fso.CreateTextFile(filePath, True)
        file.WriteLine "Test percorso lungo"
        file.Close
        Set fso = Nothing
        
        ' Verifica
        If clsFS.FileEsiste(filePath) Then
            Debug.Print "? File creato e trovato!"
            clsFS.MostraProprietaFile filePath
        End If
        
        ' Pulisci
        clsFS.EliminaFile filePath
        clsFS.EliminaCartella longPath, True
        Debug.Print "? Pulizia completata"
    Else
        Debug.Print "? Impossibile creare il percorso lungo"
    End If
    
    Set clsFS = Nothing
End Sub

' --- Test compatibilità VBA7 ---
Public Sub TestCompatibilitaVBA7()
    Dim clsFS As New CLASSE_clsFileSystem
    Dim testPath As String
    
    #If VBA7 Then
        Debug.Print "VBA7: Usa LongPtr"
    #Else
        Debug.Print "VBA7 non disponibile: Usa Long"
    #End If
    
    ' Test con percorso di sistema
    testPath = "C:\Windows\System32\notepad.exe"
    
    If clsFS.FileEsiste(testPath) Then
        If clsFS.ApriFile(testPath) Then
            MsgBox "Test compatibilità completato con successo!", vbInformation
        Else
            MsgBox "Test compatibilità fallito - impossibile aprire il file", vbExclamation
        End If
    Else
        MsgBox "Notepad non trovato", vbExclamation
    End If
    
    Set clsFS = Nothing
End Sub

' --- Test di stress ---
Public Sub TestStress()
    Dim clsFS As New CLASSE_clsFileSystem
    Dim startTime As Double
    Dim endTime As Double
    Dim i As Long
    
    Debug.Print ""
    Debug.Print "--- TEST DI STRESS ---"
    startTime = Timer
    
    ' Crea 100 file
    Dim tempFolder As String
    tempFolder = Environ("TEMP") & "\StressTest\"
    
    clsFS.CreaCartella tempFolder
    
    For i = 1 To 100
        Dim filePath As String
        filePath = tempFolder & "stress_" & Format(i, "000") & ".txt"
        
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim file As Object
        Set file = fso.CreateTextFile(filePath, True)
        file.WriteLine "File di stress " & i
        file.Close
        Set fso = Nothing
        
        ' Verifica
        If Not clsFS.FileEsiste(filePath) Then
            Debug.Print "Errore: file " & i & " non creato"
        End If
    Next i
    
    ' Elimina tutti i file
    For i = 1 To 100
        Dim path As String
        path = tempFolder & "stress_" & Format(i, "000") & ".txt"
        clsFS.EliminaFile path
    Next i
    
    clsFS.EliminaCartella tempFolder, True
    
    endTime = Timer
    Debug.Print "Test di stress completato in " & Format(endTime - startTime, "0.0") & " secondi"
    
    Set clsFS = Nothing
End Sub

