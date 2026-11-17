# TABELLA Msys_TB99_LINK.MD = TABELLA DI TUTTI I COLLEGAMENTI DEL PROGETTO

posizione: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_TB99_LINK\MDB\
## ATTIVITA
la tabella TB99 link contiene tutte le tabelle collegate del progetto con le caratteristiche,
ossia tabelle collegate esempio:
   TABELLA              TYPE     CONNECT
   MSys_DF01_COMANDI     6       ;DATABASE=C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_DF01\MDB\MSys_DF01.mdb

occorre trasformare il campo database in path quindi contenere solo la path 
il tipo di oggetto ho creato questo codice TABLES_LINK per distinguerla da table come tabella vera

il codic progetto MSys_TB99_LINK per identificare dove si trovano le informazioni 

Lo scopo di questo progetto è quello di concentrare tutte le tabelle link in tb99 con 
la tabella Msys_TB99_LINK 
   es. MSys_DF01_COMANDI

Name                 Type  Connect                                                                       Database
MSys_DF01_COMANDI 6  ;DATABASE=C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_DF01\MDB\MSys_DF01.mdb    C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_DF01\MDB\MSys_DF01.mdb

avere la path completa di dove si trova la tabella fisica che viene collegata e che viene espressa
in connett, ed avere altri due campi per utilizzi futuri ossia i Type = 6 = tabella collegata
Connect = il database completo 
Database = QUI E' DA CORREGGERE INSERIRE SOLO IL DATABASE SENZA PATH 
Database_s = QUI SOLO LA PATH SENZA IL DATABASE 
   c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb
   C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF01\MDB\MSys_DF01.mdb

lo scopo quindi della Tb99 è quello di inviduare dove si trova la tabella fisica e quindi quali 
altri oggetti sono collegati ad essa Query, Form ec...

quindi la Msys_TB99_LINK
            |
            |------>Msys_TB99_LINK_Object
è collegata alla link obiect per campo name in queto modo posso avere i quadro completo degli 
oggetti che sono collegati alla tabella e dove si trova il db originale.

La tabella Msys_TB99_LINK_Object
            |
            |------>Msys_TB99_LINK_PATH_DB
si dovra collegare alla tb99 link path db = per indicare in quale database devo esportare
la tabella collegata, gli oggetti ad essi appartenenti al microprogetto originale.
Per cui Msys_TB99_LINK_PATH_DB puo rappresentare un oggetto es. 

Perche ho costruito questo progetto? Perche devo avere la liberata e la precisione di spostare
gli oggetti dal progetto originale ad un'alro progetto e spostare anche il link della tabella per farlo
funzionare. Es. nel PROGETTO MASTER  MSYS_ATTIVA_GEST_OGGETTI.mdb ho bisogno 
dei menu per cui con la tabella TB99 link posso spostare il link della tabella GEST_MENU_Tb01_VOCI
ed eventualmente aprire il progetto tb01 voci e spostare le query dei menu tipo
   GEST_MENU_Tb01Qry01_01_VOCI
   GEST_MENU_Tb01Qry01_01_VOCI_TMP
      ecc.

Tutti gli oggetti dovrebbero essere salvati nel front end GE_MSys_TB06_MENU.mdb in modo 
da poter esportare gli oggetti query, form ecc dopo l'aggiornamento da
   GE_MSys_TB06_MENU.mdb ---->  MSYS_ATTIVA_GEST_OGGETTI.mdb

Ora mi devi riorganizzare il back end del tb99 delle tre tabelle in modo che 
      Msys_TB99_LINK             = abbia i campi essenziali per individuare l'oggetto tabella e dove si trova;
      Msys_TB99_LINK_Object      = abbia i campi essenziali per rapresentare tutti gli oggetti che devono funzionare con 
                                   la tabella sopra indicata

     Msys_TB99_LINK_PATH_DB      = i dabatabase in cui esportare gli oggetti.

Semplifica il numero dei campi collega le 3 tabelle con i campi chiavi e sistema tutte le query.
Puoi farlo?
