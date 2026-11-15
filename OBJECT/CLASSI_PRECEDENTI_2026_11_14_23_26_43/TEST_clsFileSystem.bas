Attribute VB_Name = "TEST_clsFileSystem"
' =====================================================
' MODULO: modCreaSottocartella260
' DESCRIZIONE: Crea una sottocartella di 260 caratteri sotto C:\TEMP\
' =====================================================
Option Compare Database
Option Explicit

Private fs As CLASSE_clsFileSystem
Private Const MAX_PATH As Long = 260

' =====================================================
' CREA SOTTOCARTELLA DI 260 CARATTERI
' =====================================================
Public Sub CreaSottocartella260()
    Set fs = New CLASSE_clsFileSystem
    
    Debug.Print "========================================"
    Debug.Print "CREAZIONE SOTTOCARTELLA DI 260 CARATTERI"
    Debug.Print "========================================"
    Debug.Print ""
    
    ' 1. Percorso base C:\TEMP\
    Dim basePath As String
    basePath = "C:\TEMP\"
    
    Debug.Print "Percorso base: " & basePath
    Debug.Print "Lunghezza base: " & Len(basePath) & " caratteri"
    Debug.Print ""
    
    ' 2. Verifica che C:\TEMP\ esista
    If Not fs.CartellaEsiste(basePath) Then
        Debug.Print "Creazione di C:\TEMP\..."
        fs.CreaCartella basePath
        If Not fs.CartellaEsiste(basePath) Then
            MsgBox "Impossibile creare C:\TEMP\", vbCritical
            Set fs = Nothing
            Exit Sub
        End If
    End If
    
    ' 3. Crea una cartella "Test260" sotto C:\TEMP\
    Dim cartellaMadre As String
    cartellaMadre = basePath & "Test260\"
    
    If Not fs.CartellaEsiste(cartellaMadre) Then
        fs.CreaCartella cartellaMadre
        Debug.Print "? Cartella madre creata: " & cartellaMadre
    Else
        Debug.Print "? Cartella madre già esistente: " & cartellaMadre
    End If
    Debug.Print ""
    
    ' 4. Calcola la lunghezza del nome della sottocartella
    '    Deve essere: 260 - Len(cartellaMadre) - 1 (per il backslash finale)
    Dim lunghezzaNome As Long
    lunghezzaNome = 260 - Len(cartellaMadre) - 1
    
    Debug.Print "Lunghezza nome sottocartella: " & lunghezzaNome & " caratteri"
    Debug.Print ""
    
    ' 5. Crea il nome della sottocartella
    Dim nomeSottocartella As String
    nomeSottocartella = "Sub_" & String(lunghezzaNome - 4, "X")
    
    Debug.Print "Nome sottocartella (primi 50 caratteri): " & Left(nomeSottocartella, 50) & "..."
    Debug.Print "Lunghezza nome: " & Len(nomeSottocartella) & " caratteri"
    Debug.Print ""
    
    ' 6. Costruisci il percorso completo
    Dim percorsoCompleto As String
    percorsoCompleto = cartellaMadre & nomeSottocartella & "\"
    
    Debug.Print "PERCORSO COMPLETO (" & Len(percorsoCompleto) & " caratteri):"
    Debug.Print percorsoCompleto
    Debug.Print ""
    
    ' 7. Verifica che sia esattamente 260 caratteri
    If Len(percorsoCompleto) = 260 Then
        Debug.Print "? LUNGHEZZA CORRETTA: 260 caratteri"
    Else
        Debug.Print "? ERRORE: Lunghezza " & Len(percorsoCompleto) & " caratteri (dovrebbe essere 260)"
        Set fs = Nothing
        Exit Sub
    End If
    
    ' 8. TENTATIVO DI CREAZIONE
    Debug.Print ""
    Debug.Print "--- TENTATIVO DI CREAZIONE ---"
    
    ' Metodo 1: Usa la classe
    On Error Resume Next
    Dim risultato As Boolean
    risultato = fs.CreaCartella(percorsoCompleto)
    Dim erroreNum As Long
    erroreNum = Err.Number
    Dim erroreDesc As String
    erroreDesc = Err.Description
    On Error GoTo 0
    
    If risultato Then
        Debug.Print "? SOTTOCARTELLA CREATA CON SUCCESSO!"
        Debug.Print "  " & percorsoCompleto
        
        ' Verifica
        If fs.CartellaEsiste(percorsoCompleto) Then
            Debug.Print "? Verifica esistenza: OK"
            
            ' Crea file di test
            CreaFileTestSottocartella percorsoCompleto
        End If
    Else
        Debug.Print "? ERRORE: Impossibile creare con il metodo standard"
        Debug.Print "  Errore: " & erroreNum & " - " & erroreDesc
        Debug.Print ""
        
        ' Metodo 2: Usa il prefisso \\?\ per percorsi lunghi
        Debug.Print "--- TENTATIVO CON PREFISSO \\?\ ---"
        Dim percorsoConPrefisso As String
        percorsoConPrefisso = "\\?\" & percorsoCompleto
        
        On Error Resume Next
        risultato = fs.CreaCartella(percorsoConPrefisso)
        erroreNum = Err.Number
        erroreDesc = Err.Description
        On Error GoTo 0
        
        If risultato Then
            Debug.Print "? SOTTOCARTELLA CREATA CON PREFISSO \\?\"
            Debug.Print "  " & percorsoConPrefisso
            
            If fs.CartellaEsiste(percorsoCompleto) Then
                Debug.Print "? Verifica esistenza: OK"
                CreaFileTestSottocartella percorsoCompleto
            End If
        Else
            Debug.Print "? ERRORE: Impossibile creare anche con prefisso"
            Debug.Print "  Errore: " & erroreNum & " - " & erroreDesc
            Debug.Print ""
            
            ' Metodo 3: Crea livello per livello
            Debug.Print "--- TENTATIVO CON CREAZIONE LIVELLO PER LIVELLO ---"
            CreaSottocartellaPassoPasso percorsoCompleto
        End If
    End If
    
    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "OPERAZIONE COMPLETATA"
    Debug.Print "========================================"
    
    Set fs = Nothing
End Sub

' =====================================================
' CREA FILE DI TEST NELLA SOTTOCARTELLA
' =====================================================
Private Sub CreaFileTestSottocartella(percorso As String)
    On Error Resume Next
    
    Dim filePath As String
    filePath = percorso & "test_sottocartella.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.CreateTextFile(filePath, True)
    
    If Not file Is Nothing Then
        file.WriteLine "========================================"
        file.WriteLine "TEST SOTTOCARTELLA DI 260 CARATTERI"
        file.WriteLine "========================================"
        file.WriteLine ""
        file.WriteLine "Percorso completo: " & percorso
        file.WriteLine "Lunghezza percorso: " & Len(percorso) & " caratteri"
        file.WriteLine "Creato il: " & Now
        file.WriteLine ""
        file.WriteLine "Questa è una sottocartella di 260 caratteri"
        file.WriteLine "creata sotto C:\TEMP\Test260\"
        file.WriteLine "========================================"
        file.Close
        
        Debug.Print "? File di test creato: " & filePath
        
        ' Mostra proprietà
        fs.MostraProprietaFile filePath
    Else
        Debug.Print "? ERRORE: Impossibile creare il file di test"
    End If
    
    Set fso = Nothing
    On Error GoTo 0
End Sub

' =====================================================
' CREAZIONE LIVELLO PER LIVELLO
' =====================================================
Private Sub CreaSottocartellaPassoPasso(percorsoCompleto As String)
    Dim parti() As String
    Dim percorsoCorrente As String
    Dim i As Long
    Dim risultato As Boolean
    
    ' Dividi il percorso
    parti = Split(percorsoCompleto, "\")
    
    ' Ricostruisci livello per livello
    For i = LBound(parti) To UBound(parti) - 1
        If i = 0 Then
            percorsoCorrente = parti(i)
        Else
            percorsoCorrente = percorsoCorrente & "\" & parti(i)
            
            If Not fs.CartellaEsiste(percorsoCorrente & "\") Then
                On Error Resume Next
                risultato = fs.CreaCartella(percorsoCorrente & "\")
                On Error GoTo 0
                
                If risultato Then
                    Debug.Print "? Creato: " & percorsoCorrente & "\"
                Else
                    Debug.Print "? ERRORE al livello " & i & ": " & percorsoCorrente & "\"
                    Exit Sub
                End If
            Else
                Debug.Print "? Esiste già: " & percorsoCorrente & "\"
            End If
        End If
    Next i
    
    ' Verifica finale
    If fs.CartellaEsiste(percorsoCompleto) Then
        Debug.Print "? SOTTOCARTELLA CREATA CON SUCCESSO!"
        CreaFileTestSottocartella percorsoCompleto
    End If
End Sub

' =====================================================
' CREA SOTTOCARTELLA CON LUNGHEZZA PERSONALIZZATA
' =====================================================
Public Sub CreaSottocartellaPersonalizzata()
    Dim fsLocal As New CLASSE_clsFileSystem
    Dim basePath As String
    Dim cartellaMadre As String
    Dim lunghezzaDesiderata As Long
    Dim prefisso As String
    Dim nome As String
    
    basePath = "C:\TEMP\"
    
    ' Crea cartella madre se non esiste
    cartellaMadre = basePath & "Test260\"
    If Not fsLocal.CartellaEsiste(cartellaMadre) Then
        fsLocal.CreaCartella cartellaMadre
    End If
    
    ' Chiedi la lunghezza desiderata
    lunghezzaDesiderata = InputBox( _
        "Inserisci la lunghezza desiderata (260-300):", _
        "Lunghezza Sottocartella", 260)
    
    If lunghezzaDesiderata < 260 Or lunghezzaDesiderata > 300 Then
        MsgBox "Inserisci un valore tra 260 e 300", vbExclamation
        Set fsLocal = Nothing
        Exit Sub
    End If
    
    ' Chiedi un prefisso
    prefisso = InputBox("Inserisci un prefisso (es. TEST_, MY_, DATA_):", "Prefisso", "SUB_")
    
    ' Calcola la lunghezza del nome
    Dim lunghezzaNome As Long
    lunghezzaNome = lunghezzaDesiderata - Len(cartellaMadre) - Len(prefisso) - 1
    
    ' Crea il nome
    nome = prefisso & String(lunghezzaNome, "X")
    
    ' Costruisci il percorso
    Dim percorso As String
    percorso = cartellaMadre & nome & "\"
    
    Debug.Print "Lunghezza percorso: " & Len(percorso) & " caratteri"
    Debug.Print "Percorso: " & Left(percorso, 80) & "..."
    
    ' Crea
    If fsLocal.CreaCartella(percorso) Then
        Debug.Print "? SOTTOCARTELLA CREATA!"
        MsgBox "Sottocartella creata con successo!" & vbCrLf & _
               "Lunghezza: " & Len(percorso) & " caratteri", vbInformation
    Else
        Debug.Print "? ERRORE: Creazione fallita"
        MsgBox "Impossibile creare la sottocartella", vbCritical
    End If
    
    Set fsLocal = Nothing
End Sub

' =====================================================
' TEST CON FILE IN SOTTOCARTELLA
' =====================================================
Public Sub TestFileInSottocartella260()
    Dim fsLocal As New CLASSE_clsFileSystem
    
    ' Crea la sottocartella di 260 caratteri
    CreaSottocartella260
    
    ' Cerca la sottocartella creata
    Dim basePath As String
    basePath = "C:\TEMP\Test260\"
    
    If fsLocal.CartellaEsiste(basePath) Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim folder As Object
        Set folder = fso.GetFolder(basePath)
        
        Dim subfolder As Object
        For Each subfolder In folder.SubFolders
            If Len(subfolder.path) = 260 Then
                Debug.Print ""
                Debug.Print "========================================"
                Debug.Print "SOTTOCARTELLA TROVATA:"
                Debug.Print "========================================"
                Debug.Print "Percorso: " & subfolder.path
                Debug.Print "Lunghezza: " & Len(subfolder.path) & " caratteri"
                Debug.Print ""
                
                ' Crea un file nella sottocartella
                Dim filePath As String
                filePath = subfolder.path & "\test_finale.txt"
                
                Dim file As Object
                Set file = fso.CreateTextFile(filePath, True)
                file.WriteLine "Test completato!"
                file.WriteLine "Percorso di " & Len(subfolder.path) & " caratteri"
                file.Close
                
                If fsLocal.FileEsiste(filePath) Then
                    Debug.Print "? File creato con successo!"
                    fsLocal.MostraProprietaFile filePath
                End If
                
                Exit For
            End If
        Next subfolder
        
        Set fso = Nothing
    End If
    
    Set fsLocal = Nothing
End Sub

' =====================================================
' PULIZIA
' =====================================================
Public Sub PuliziaSottocartella260()
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim contatore As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    contatore = 0
    
    Debug.Print "--- PULIZIA SOTTOCARTELLE DI TEST ---"
    
    ' Cancella la cartella Test260 e tutto il suo contenuto
    Dim testPath As String
    testPath = "C:\TEMP\Test260\"
    
    If fso.FolderExists(testPath) Then
        On Error Resume Next
        fso.DeleteFolder testPath, True
        If Err.Number = 0 Then
            Debug.Print "? Cartella Test260 cancellata: " & testPath
            contatore = 1
        Else
            Debug.Print "? Errore: " & Err.Description
        End If
        On Error GoTo 0
    Else
        Debug.Print "Nessuna cartella Test260 trovata."
    End If
    
    ' Cerca anche altre cartelle di test in C:\TEMP\
    If fso.FolderExists("C:\TEMP\") Then
        Set folder = fso.GetFolder("C:\TEMP\")
        
        For Each subfolder In folder.SubFolders
            Dim nome As String
            nome = subfolder.Name
            
            ' Cerca cartelle di test (iniziano con Sub_, SUB_, TEST_, MY_, DATA_)
            If Left(nome, 4) = "Sub_" Or Left(nome, 4) = "SUB_" Or _
               Left(nome, 5) = "TEST_" Or Left(nome, 3) = "MY_" Or Left(nome, 5) = "DATA_" Then
                
                If Len(nome) > 100 Then
                    contatore = contatore + 1
                    Debug.Print "Trovata: " & subfolder.path & " (" & Len(subfolder.path) & " caratteri)"
                    
                    On Error Resume Next
                    fso.DeleteFolder subfolder.path, True
                    If Err.Number = 0 Then
                        Debug.Print "  ? Cancellata"
                    Else
                        Debug.Print "  ? Errore: " & Err.Description
                    End If
                    On Error GoTo 0
                End If
            End If
        Next subfolder
    End If
    
    If contatore = 0 Then
        Debug.Print "Nessuna cartella di test trovata."
    Else
        Debug.Print "Totale cartelle cancellate: " & contatore
    End If
    
    Set fso = Nothing
End Sub

