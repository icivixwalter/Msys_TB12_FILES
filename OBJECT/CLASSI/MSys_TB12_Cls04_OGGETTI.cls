VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSys_TB12_Cls04_OGGETTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'==================================================================================================
'  CLASSE    : MSys_TB12_Cls04_OGGETTI
'  PROGETTO  : MSys_TB12_FILES  (mini-progetto TB12)
'  SCOPO     : oggetti di Access tra database: elenco (tabelle, query, form, report, macro, moduli)
'              del database corrente o di uno esterno, esistenza, importazione ed esportazione di
'              singoli oggetti o di gruppi (filtro sul nome), registrazione dell'elenco in una tabella
'              di inventario (predefinita: Msys_TB03_ELENCO_ObjProgetti) senza doppioni.
'  SOSTITUISCE: Msys_TB12_FILES_CLASSE_clsDBInspector e CLASSE_CLASSE_clsDBInspector
'              (quelle eseguivano 3 query di accodamento che nel database non esistono)
'  TIPI      : TABLE, QUERY, FORMS, REPORTS, MACRO, MODULI (come in MSys_DF02_TIPO_OGGETTO) e TUTTI
'  DIPENDENZE: riferimento DAO. Nessuna classe.
'  LOG       : <cartella del database>\Logs\MSys_TB12_Cls04_OGGETTI_AAAA_MM_GG.txt
'  VERSIONE  : 1.0 - 2026-11-15
'--------------------------------------------------------------------------------------------------
'  INDICE DELLE SEZIONI, DEI PARAGRAFI E DEI SOTTOPARAGRAFI  (cerca l'etichetta, es. #05.03.02#)
'--------------------------------------------------------------------------------------------------
'  #01#  IMPOSTAZIONI = costanti dei tipi, tabella di inventario predefinita, variabili
'        #01.01# = costanti          #01.02# = variabili della classe
'  #02#  CICLO_DI_VITA_E_PROPRIETA = inizializzazione e impostazioni
'        #02.01# = Class_Initialize  #02.02# = Class_Terminate
'        #02.03# = proprieta: TabellaElenco, CampoProgetto, CampoTipo, CampoNome, CampoData, log
'  #03#  ELENCO = oggetti di un database
'        #03.01# = ElencoOggetti     #03.01.01# apertura db  #03.01.02# tabelle e query  #03.01.03# contenitori  #03.01.04# chiusura
'        #03.02# = ContaOggetti
'  #04#  ESISTENZA
'        #04.01# = OggettoEsiste
'  #05#  TRASFERIMENTO = importazione ed esportazione con DoCmd.TransferDatabase
'        #05.01# = ImportaOggetto    #05.01.01# controlli  #05.01.02# esistente  #05.01.03# importazione
'        #05.02# = EsportaOggetto    #05.02.01# controlli  #05.02.02# esportazione
'        #05.03# = AccodaOggetti     #05.03.01# elenco sorgente  #05.03.02# importazione una per una  #05.03.03# esito
'  #06#  INVENTARIO = registrazione dell'elenco in una tabella
'        #06.01# = AccodaElencoInTabella #06.01.01# tabella e campi  #06.01.02# elenco  #06.01.03# inserimento senza doppioni
'        #06.02# = AccodaTuttiGliOggetti (compatibile: tabelle, query e form del database corrente)
'  #07#  CONVERSIONI = tipo testo <-> costante di Access
'        #07.01# = TipoAccess        #07.02# = TipoValido
'  #90#  LOG_INTERNO = LOG_Scrivi, LOG_Errore, LOG_Passo, LOG_Cartella
'==================================================================================================


'##################################################################################################
'## #01# IMPOSTAZIONI = costanti dei tipi, tabella di inventario predefinita, variabili
'##################################################################################################

'//#01.01# = costanti
Private Const NOME_CLASSE As String = "MSys_TB12_Cls04_OGGETTI"
Private Const TIPI_TUTTI As String = "TABLE|QUERY|FORMS|REPORTS|MACRO|MODULI"

'//#01.02# = variabili della classe
Private m_TabellaElenco_s As String
Private m_CampoProgetto_s As String
Private m_CampoTipo_s As String
Private m_CampoNome_s As String
Private m_CampoData_s As String
Private m_CartellaLog_s As String
Private m_LogAttivo_b As Boolean
Private m_LogDettagliato_b As Boolean
Private m_UltimoErrore_s As String


'##################################################################################################
'## #02# CICLO_DI_VITA_E_PROPRIETA = inizializzazione e impostazioni
'##################################################################################################

Private Sub Class_Initialize()
    '//#02.01# = inventario predefinito: Msys_TB03_ELENCO_ObjProgetti (campi del mini-progetto TB03)
    m_TabellaElenco_s = "Msys_TB03_ELENCO_ObjProgetti"
    m_CampoProgetto_s = "COD_PROGETTO_s"
    m_CampoTipo_s = "TIPOGGETTO_s"
    m_CampoNome_s = "NOMEOGGETTO_s"
    m_CampoData_s = "DATAINS_d"
    m_LogAttivo_b = True
    m_LogDettagliato_b = False
End Sub

Private Sub Class_Terminate()
    '//#02.02# = nessuna risorsa aperta: i database esterni si chiudono in ogni metodo
End Sub

'//#02.03# = proprieta
Public Property Get TabellaElenco() As String
    TabellaElenco = m_TabellaElenco_s
End Property
Public Property Let TabellaElenco(ByVal v As String)
    m_TabellaElenco_s = v
End Property
Public Property Get CampoProgetto() As String
    CampoProgetto = m_CampoProgetto_s
End Property
Public Property Let CampoProgetto(ByVal v As String)
    m_CampoProgetto_s = v
End Property
Public Property Get CampoTipo() As String
    CampoTipo = m_CampoTipo_s
End Property
Public Property Let CampoTipo(ByVal v As String)
    m_CampoTipo_s = v
End Property
Public Property Get CampoNome() As String
    CampoNome = m_CampoNome_s
End Property
Public Property Let CampoNome(ByVal v As String)
    m_CampoNome_s = v
End Property
Public Property Get CampoData() As String
    CampoData = m_CampoData_s
End Property
Public Property Let CampoData(ByVal v As String)
    '//stringa vuota = la data non viene scritta
    m_CampoData_s = v
End Property
Public Property Get CartellaLog() As String
    CartellaLog = LOG_Cartella()
End Property
Public Property Let CartellaLog(ByVal v As String)
    m_CartellaLog_s = v
End Property
Public Property Get LogAttivo() As Boolean
    LogAttivo = m_LogAttivo_b
End Property
Public Property Let LogAttivo(ByVal v As Boolean)
    m_LogAttivo_b = v
End Property
Public Property Get LogDettagliato() As Boolean
    LogDettagliato = m_LogDettagliato_b
End Property
Public Property Let LogDettagliato(ByVal v As Boolean)
    m_LogDettagliato_b = v
End Property
Public Property Get UltimoErrore() As String
    UltimoErrore = m_UltimoErrore_s
End Property


'##################################################################################################
'## #03# ELENCO = oggetti di un database
'##################################################################################################

Public Function ElencoOggetti(Optional ByVal tipo As String = "TUTTI", Optional ByVal filtroNome As String = "*", _
                              Optional ByVal percorsoDb As String = vbNullString, Optional ByVal escludiSistema As Boolean = True) As Collection
'//#03.01# = Collection di "TIPO|NOME" del database corrente (percorsoDb vuoto) o di un database esterno
Const PROC As String = "ElencoOggetti"
Dim sPos_s As String, coll As New Collection, db As DAO.database, bEsterno As Boolean
Dim td As DAO.TableDef, qd As DAO.QueryDef, doc As DAO.Document, cnt As DAO.Container, vT As Variant, sCont_s As String, sTipi_s As String
On Error GoTo ERRORE
    '//#03.01.01# = apertura: corrente oppure esterno in sola lettura
    sPos_s = "03.01.01"
    tipo = UCase$(tipo)
    If tipo = "TUTTI" Then sTipi_s = TIPI_TUTTI Else sTipi_s = tipo
    If Len(percorsoDb) = 0 Then
        Set db = CurrentDb
    Else
        Set db = DBEngine.OpenDatabase(percorsoDb, False, True)
        bEsterno = True
    End If
    For Each vT In Split(sTipi_s, "|")
        Select Case vT
            Case "TABLE"
                '//#03.01.02# = tabelle (esclusi MSys* e ~* se escludiSistema)
                sPos_s = "03.01.02"
                For Each td In db.TableDefs
                    If Not (escludiSistema And (td.Name Like "MSys*" Or td.Name Like "~*")) Then
                        If td.Name Like filtroNome Then coll.Add "TABLE|" & td.Name
                    End If
                Next td
            Case "QUERY"
                sPos_s = "03.01.02"
                For Each qd In db.QueryDefs
                    If Not (escludiSistema And qd.Name Like "~*") Then
                        If qd.Name Like filtroNome Then coll.Add "QUERY|" & qd.Name
                    End If
                Next qd
            Case "FORMS", "REPORTS", "MACRO", "MODULI"
                '//#03.01.03# = contenitori Forms, Reports, Scripts (macro), Modules (moduli e classi)
                sPos_s = "03.01.03"
                Select Case vT
                    Case "FORMS":   sCont_s = "Forms"
                    Case "REPORTS": sCont_s = "Reports"
                    Case "MACRO":   sCont_s = "Scripts"
                    Case Else:      sCont_s = "Modules"
                End Select
                '    un database creato con DAO e mai aperto da Access non ha ancora questi contenitori
                Set cnt = Nothing
                On Error Resume Next
                Set cnt = db.Containers(sCont_s)
                On Error GoTo ERRORE
                If Not cnt Is Nothing Then
                    For Each doc In cnt.Documents
                        If Not (escludiSistema And doc.Name Like "~*") Then
                            If doc.Name Like filtroNome Then coll.Add vT & "|" & doc.Name
                        End If
                    Next doc
                End If
            Case Else
                LOG_Scrivi PROC, sPos_s, "AVVISO", "tipo non previsto: " & vT
        End Select
    Next vT
    LOG_Passo PROC, sPos_s, coll.Count & " oggetti " & tipo & " " & filtroNome & " in " & IIf(bEsterno, percorsoDb, "database corrente")
USCITA:
    '//#03.01.04# = chiusura del database esterno
    On Error Resume Next
    If bEsterno Then db.Close
    Set db = Nothing
    Set ElencoOggetti = coll
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & percorsoDb & "]"
    Resume USCITA
End Function

Public Function ContaOggetti(Optional ByVal tipo As String = "TUTTI", Optional ByVal filtroNome As String = "*", Optional ByVal percorsoDb As String = vbNullString) As Long
'//#03.02#
    ContaOggetti = ElencoOggetti(tipo, filtroNome, percorsoDb).Count
End Function


'##################################################################################################
'## #04# ESISTENZA
'##################################################################################################

Public Function OggettoEsiste(ByVal nome As String, ByVal tipo As String, Optional ByVal percorsoDb As String = vbNullString) As Boolean
'//#04.01# = True se l'oggetto del tipo indicato esiste (confronto esatto: i caratteri # ? [ del nome non fanno da filtro)
Dim v As Variant
    For Each v In ElencoOggetti(tipo, "*", percorsoDb, False)
        If StrComp(Mid$(v, InStr(v, "|") + 1), nome, vbTextCompare) = 0 Then OggettoEsiste = True: Exit Function
    Next v
End Function


'##################################################################################################
'## #05# TRASFERIMENTO = importazione ed esportazione con DoCmd.TransferDatabase
'##################################################################################################

Public Function ImportaOggetto(ByVal percorsoSorgente As String, ByVal tipo As String, ByVal nome As String, _
                               Optional ByVal nomeDestinazione As String = vbNullString, Optional ByVal sovrascrivi As Boolean = False) As Boolean
'//#05.01# = copia un oggetto da un altro database nel database corrente
Const PROC As String = "ImportaOggetto"
Dim sPos_s As String
On Error GoTo ERRORE
    '//#05.01.01# = controlli: tipo valido, file sorgente, oggetto presente nella sorgente
    sPos_s = "05.01.01"
    tipo = UCase$(tipo)
    If Len(nomeDestinazione) = 0 Then nomeDestinazione = nome
    If Not TipoValido(tipo) Then LOG_Scrivi PROC, sPos_s, "ERRORE", "tipo non valido: " & tipo: GoTo USCITA
    If Len(Dir$(percorsoSorgente)) = 0 Then LOG_Scrivi PROC, sPos_s, "ERRORE", "database sorgente inesistente: " & percorsoSorgente: GoTo USCITA
    If Not OggettoEsiste(nome, tipo, percorsoSorgente) Then LOG_Scrivi PROC, sPos_s, "ERRORE", tipo & " " & nome & " non presente in " & percorsoSorgente: GoTo USCITA
    '//#05.01.02# = oggetto gia' presente nel database corrente: salto oppure elimino (sovrascrivi)
    sPos_s = "05.01.02"
    If OggettoEsiste(nomeDestinazione, tipo) Then
        If Not sovrascrivi Then
            LOG_Scrivi PROC, sPos_s, "AVVISO", tipo & " " & nomeDestinazione & " esiste gia': non importato"
            GoTo USCITA
        End If
        DoCmd.DeleteObject TipoAccess(tipo), nomeDestinazione
        LOG_Scrivi PROC, sPos_s, "INFO", tipo & " " & nomeDestinazione & " eliminato per sovrascrittura"
    End If
    '//#05.01.03# = importazione (tabelle con i dati)
    sPos_s = "05.01.03"
    DoCmd.TransferDatabase acImport, "Microsoft Access", percorsoSorgente, TipoAccess(tipo), nome, nomeDestinazione, False
    ImportaOggetto = True
    LOG_Scrivi PROC, sPos_s, "OK", tipo & " " & nome & IIf(nome <> nomeDestinazione, " -> " & nomeDestinazione, "") & " da " & percorsoSorgente
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & tipo & " " & nome & "]"
    ImportaOggetto = False
    Resume USCITA
End Function

Public Function EsportaOggetto(ByVal percorsoDestinazione As String, ByVal tipo As String, ByVal nome As String, _
                               Optional ByVal nomeDestinazione As String = vbNullString) As Boolean
'//#05.02# = copia un oggetto del database corrente in un altro database (che non deve essere aperto in modo esclusivo)
Const PROC As String = "EsportaOggetto"
Dim sPos_s As String
On Error GoTo ERRORE
    '//#05.02.01# = controlli
    sPos_s = "05.02.01"
    tipo = UCase$(tipo)
    If Len(nomeDestinazione) = 0 Then nomeDestinazione = nome
    If Not TipoValido(tipo) Then LOG_Scrivi PROC, sPos_s, "ERRORE", "tipo non valido: " & tipo: GoTo USCITA
    If Len(Dir$(percorsoDestinazione)) = 0 Then LOG_Scrivi PROC, sPos_s, "ERRORE", "database di destinazione inesistente: " & percorsoDestinazione: GoTo USCITA
    If Not OggettoEsiste(nome, tipo) Then LOG_Scrivi PROC, sPos_s, "ERRORE", tipo & " " & nome & " non presente nel database corrente": GoTo USCITA
    '//#05.02.02# = esportazione (se esiste nella destinazione viene sostituito)
    sPos_s = "05.02.02"
    DoCmd.TransferDatabase acExport, "Microsoft Access", percorsoDestinazione, TipoAccess(tipo), nome, nomeDestinazione, False
    EsportaOggetto = True
    LOG_Scrivi PROC, sPos_s, "OK", tipo & " " & nome & IIf(nome <> nomeDestinazione, " -> " & nomeDestinazione, "") & " in " & percorsoDestinazione
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & tipo & " " & nome & "]"
    EsportaOggetto = False
    Resume USCITA
End Function

Public Function AccodaOggetti(ByVal percorsoSorgente As String, Optional ByVal tipo As String = "TUTTI", _
                              Optional ByVal filtroNome As String = "*", Optional ByVal sovrascrivi As Boolean = False) As Long
'//#05.03# = importa nel database corrente tutti gli oggetti della sorgente che corrispondono a tipo e filtro.
'           Restituisce quanti oggetti sono stati importati; -1 in caso di errore
Const PROC As String = "AccodaOggetti"
Dim sPos_s As String, v As Variant, vP As Variant, nOk_l As Long, nNo_l As Long, coll As Collection
On Error GoTo ERRORE
    '//#05.03.01# = elenco degli oggetti nella sorgente
    sPos_s = "05.03.01"
    Set coll = ElencoOggetti(tipo, filtroNome, percorsoSorgente)
    LOG_Scrivi PROC, sPos_s, "INIZIO", coll.Count & " oggetti " & UCase$(tipo) & " " & filtroNome & " da " & percorsoSorgente
    '//#05.03.02# = importazione uno per uno: un errore non ferma gli altri
    sPos_s = "05.03.02"
    For Each v In coll
        vP = Split(v, "|")
        If ImportaOggetto(percorsoSorgente, vP(0), vP(1), vbNullString, sovrascrivi) Then nOk_l = nOk_l + 1 Else nNo_l = nNo_l + 1
    Next v
    '//#05.03.03# = esito
    sPos_s = "05.03.03"
    AccodaOggetti = nOk_l
    LOG_Scrivi PROC, sPos_s, IIf(nNo_l = 0, "OK", "AVVISO"), "importati " & nOk_l & ", non importati " & nNo_l
USCITA:
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description
    AccodaOggetti = -1
    Resume USCITA
End Function


'##################################################################################################
'## #06# INVENTARIO = registrazione dell'elenco in una tabella
'##################################################################################################

Public Function AccodaElencoInTabella(ByVal codiceProgetto As String, Optional ByVal percorsoDb As String = vbNullString, _
                                      Optional ByVal tipo As String = "TUTTI", Optional ByVal filtroNome As String = "*") As Long
'//#06.01# = scrive gli oggetti in TabellaElenco (locale o collegata) per il progetto indicato, senza doppioni.
'           Restituisce le righe aggiunte; -1 in caso di errore
Const PROC As String = "AccodaElencoInTabella"
Dim sPos_s As String, db As DAO.database, rs As DAO.Recordset, v As Variant, vP As Variant, nAgg_l As Long, nGia_l As Long, sCrit_s As String
On Error GoTo ERRORE
    '//#06.01.01# = tabella e campi presenti nel database corrente
    sPos_s = "06.01.01"
    Set db = CurrentDb
    If Not CampoPresente(db, m_TabellaElenco_s, m_CampoProgetto_s) Or Not CampoPresente(db, m_TabellaElenco_s, m_CampoTipo_s) _
       Or Not CampoPresente(db, m_TabellaElenco_s, m_CampoNome_s) Then
        LOG_Scrivi PROC, sPos_s, "ERRORE", "tabella o campi mancanti: " & m_TabellaElenco_s & " (" & m_CampoProgetto_s & ", " & m_CampoTipo_s & ", " & m_CampoNome_s & ")"
        m_UltimoErrore_s = "#" & sPos_s & "# tabella di inventario mancante: " & m_TabellaElenco_s
        AccodaElencoInTabella = -1
        GoTo USCITA
    End If
    '//#06.01.02# = elenco degli oggetti
    sPos_s = "06.01.02"
    Set rs = db.OpenRecordset(m_TabellaElenco_s, dbOpenDynaset)
    '//#06.01.03# = inserimento: salto le righe gia' presenti (progetto + tipo + nome)
    sPos_s = "06.01.03"
    For Each v In ElencoOggetti(tipo, filtroNome, percorsoDb)
        vP = Split(v, "|")
        sCrit_s = "[" & m_CampoProgetto_s & "]='" & Replace(codiceProgetto, "'", "''") & "' AND [" & m_CampoTipo_s & "]='" & vP(0) & _
                  "' AND [" & m_CampoNome_s & "]='" & Replace(vP(1), "'", "''") & "'"
        rs.FindFirst sCrit_s
        If rs.NoMatch Then
            rs.AddNew
            rs.Fields(m_CampoProgetto_s).value = codiceProgetto
            rs.Fields(m_CampoTipo_s).value = vP(0)
            rs.Fields(m_CampoNome_s).value = vP(1)
            If Len(m_CampoData_s) > 0 Then If CampoPresente(db, m_TabellaElenco_s, m_CampoData_s) Then rs.Fields(m_CampoData_s).value = Now
            rs.Update
            nAgg_l = nAgg_l + 1
        Else
            nGia_l = nGia_l + 1
        End If
    Next v
    AccodaElencoInTabella = nAgg_l
    LOG_Scrivi PROC, sPos_s, "OK", codiceProgetto & ": aggiunti " & nAgg_l & ", gia' presenti " & nGia_l & " in " & m_TabellaElenco_s
USCITA:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ERRORE:
    LOG_Errore PROC, sPos_s, Err.Number, Err.Description & " [" & m_TabellaElenco_s & "]"
    AccodaElencoInTabella = -1
    Resume USCITA
End Function

Public Function AccodaTuttiGliOggetti(ByVal codiceProgetto As String, Optional ByVal mostraReport As Boolean = True) As Boolean
'//#06.02# = compatibile con la versione precedente: tabelle, query e form del database corrente nell'inventario
Dim nT As Long, nQ As Long, nF As Long, sRep_s As String, bOk As Boolean
    nT = AccodaElencoInTabella(codiceProgetto, vbNullString, "TABLE")
    nQ = AccodaElencoInTabella(codiceProgetto, vbNullString, "QUERY")
    nF = AccodaElencoInTabella(codiceProgetto, vbNullString, "FORMS")
    bOk = (nT >= 0 And nQ >= 0 And nF >= 0)
    sRep_s = IIf(nT >= 0, "[OK] Tabelle: " & nT, "[NO] Tabelle") & vbCrLf & IIf(nQ >= 0, "[OK] Query: " & nQ, "[NO] Query") & vbCrLf & _
             IIf(nF >= 0, "[OK] Form: " & nF, "[NO] Form") & vbCrLf & "(righe aggiunte in " & m_TabellaElenco_s & ")"
    LOG_Scrivi "AccodaTuttiGliOggetti", "06.02", IIf(bOk, "OK", "ERRORE"), Replace(sRep_s, vbCrLf, " | ")
    If mostraReport Then MsgBox "ACCODAMENTO OGGETTI - progetto " & codiceProgetto & vbCrLf & vbCrLf & sRep_s, IIf(bOk, vbInformation, vbExclamation), "Accodamento oggetti"
    AccodaTuttiGliOggetti = bOk
End Function


'##################################################################################################
'## #07# CONVERSIONI = tipo testo <-> costante di Access
'##################################################################################################

Public Function TipoAccess(ByVal tipo As String) As Long
'//#07.01# = TABLE 0, QUERY 1, FORMS 2, REPORTS 3, MACRO 4, MODULI 5 (acTable ... acModule); -1 se sconosciuto
    Select Case UCase$(tipo)
        Case "TABLE":   TipoAccess = acTable
        Case "QUERY":   TipoAccess = acQuery
        Case "FORMS":   TipoAccess = acForm
        Case "REPORTS": TipoAccess = acReport
        Case "MACRO":   TipoAccess = acMacro
        Case "MODULI":  TipoAccess = acModule
        Case Else:      TipoAccess = -1
    End Select
End Function

Public Function TipoValido(ByVal tipo As String) As Boolean
'//#07.02#
    TipoValido = (TipoAccess(tipo) >= 0)
End Function

Private Function CampoPresente(ByVal db As DAO.database, ByVal sTabella As String, ByVal sCampo As String) As Boolean
    On Error Resume Next
    CampoPresente = (Len(db.TableDefs(sTabella).Fields(sCampo).Name) > 0)
End Function


'##################################################################################################
'## #90# LOG_INTERNO = <db>\Logs\MSys_TB12_Cls04_OGGETTI_AAAA_MM_GG.txt
'##################################################################################################

Private Sub LOG_Scrivi(ByVal sFunz As String, ByVal sPos As String, ByVal sEsito As String, ByVal sTesto As String)
'//#90.01# = la stessa riga in Debug.Print e nel file
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
'//#90.02#
    m_UltimoErrore_s = "#" & sPos & "# " & sFunz & " - errore " & nErr & ": " & sDescr
    LOG_Scrivi sFunz, sPos, "ERRORE", "errore " & nErr & ": " & sDescr & " -> codice fermo a #" & sPos & "#"
End Sub

Private Sub LOG_Passo(ByVal sFunz As String, ByVal sPos As String, ByVal sTesto As String)
'//#90.03# = solo con LogDettagliato = True
    If m_LogDettagliato_b Then LOG_Scrivi sFunz, sPos, "INFO", sTesto
End Sub

Private Function LOG_Cartella() As String
'//#90.04# = cartella impostata oppure <cartella del database>\Logs (creata se manca)
    On Error Resume Next
    If Len(m_CartellaLog_s) > 0 Then LOG_Cartella = m_CartellaLog_s Else LOG_Cartella = CurrentProject.path & "\Logs"
    If Len(Dir$(LOG_Cartella, vbDirectory)) = 0 Then MkDir LOG_Cartella
End Function
