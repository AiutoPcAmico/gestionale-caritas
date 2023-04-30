VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelezionaUtenzaStampa 
   Caption         =   "Seleziona Utenza per Stampa"
   ClientHeight    =   3645
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6270
   OleObjectBlob   =   "SelezionaUtenzaStampa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelezionaUtenzaStampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             SelezioneUtenzaStampa         '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '


' -------------------------------------------------
' Al click della conferma, gestisco l'esportazione
' dei dati in un PDF
' -------------------------------------------------

Private Sub ConfermaScelta_Click()
    Dim DataOdierna As String
    Dim lastIndexOfConsegne As Integer
    Dim rowToNewSheet As Integer
    Dim datiUtenza As Object
    Dim selectedUtenza As Integer
    
    selectedUtenza = UtentiComboBox.Column(1)
    rowToNewSheet = 6
    lastIndexOfConsegne = getLastRowIndex("Consegne")
    DataOdierna = Format(getTodayDate(), "dd/mm/yyyy")
    Set datiUtenza = getUtenteGeneralita(selectedUtenza)
    
    dummy = MsgBox("Esportazione delle consegne per l'utente in corso!" & vbCrLf & vbCrLf & "Al termine della operazione, verrà aperto il PDF stampabile e salvato nella cartella 'stampe'", vbInformation)
    
    'svuoto il foglio per la nuova sovrascrittura
    Worksheets("StampaConsegneUtenza").Range("A6:D100").ClearContents
    
    'setto sul foglio la data odierna
    ActiveWorkbook.Sheets("StampaConsegneUtenza").Range("B2").Value = "Consegne utenze: " & datiUtenza("Cognome") & " " & datiUtenza("Nome")
    ActiveWorkbook.Sheets("StampaConsegneUtenza").Range("B3").Value = "Aggiornato al " & DataOdierna
    
    'scorro tutta la lista delle ultime consegne e copio solo quelle dell'utente selezionato
    Dim i As Integer
    Dim viveriConsegnati As String
    Dim oggettiConsegnati As String
    Dim dataConsegna As String
    
    
    
    For i = 2 To lastIndexOfConsegne
        'verifico che la data sia uguale a quella odierna
        idUtenza = ActiveWorkbook.Sheets("Consegne").Cells(i, 1).Value
        If idUtenza = selectedUtenza Then
            
            'leggo e ricopio la data di consegna
            dataConsegna = ActiveWorkbook.Sheets("Consegne").Cells(i, 2).Value
            ActiveWorkbook.Sheets("StampaConsegneUtenza").Cells(rowToNewSheet, 1).Value = dataConsegna
            
            'leggo e ricopio i viveri
            viveriConsegnati = ActiveWorkbook.Sheets("Consegne").Cells(i, 3).Value
            ActiveWorkbook.Sheets("StampaConsegneUtenza").Cells(rowToNewSheet, 2).Value = viveriConsegnati
            
            'leggo e ricopio i beni o vestiti
            oggettiConsegnati = ActiveWorkbook.Sheets("Consegne").Cells(i, 4).Value
            ActiveWorkbook.Sheets("StampaConsegneUtenza").Cells(rowToNewSheet, 3).Value = oggettiConsegnati
            
            
            'incremento la prossima riga da scrivere
            rowToNewSheet = rowToNewSheet + 1
        End If
    Next
    
    'e formatto il foglio
    ActiveWorkbook.Sheets("StampaConsegneUtenza").Cells.EntireRow.AutoFit
    
    'ed infine lo stampo
    
    'modifico la data per renderla compatibile a windows
    Dim dataModificata As String
    dataModificata = Replace(DataOdierna, "/", "-")
    
    'esporto e salvo
    ActiveWorkbook.Sheets("StampaConsegneUtenza").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ThisWorkbook.Path & "\stampe\consegne_per_utenza\" & datiUtenza("Cognome") & " " & datiUtenza("Nome") & " al " & dataModificata & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
        
    'riapro il precdente menu
    Me.Hide
    Menuattivita.Show
        
End Sub

Private Sub TornaIndietro_Click()
    Me.Hide
    Menuattivita.Show
End Sub


' -------------------------------------------------
' Al load della UserForm, carico le utenze possibili
' -------------------------------------------------


Private Sub UserForm_Activate()
 'load all users with their ID
    Dim lastRow As Integer
    Dim index As Integer
    Dim user As Object
    Dim UserID As Integer
    lastRow = getLastRowIndex("Utenti")
    Dim listOfUser() As Variant
    Dim indexOfList As Integer
    
    indexOfList = 0
    
    ReDim listOfUser(lastRow - 2, 2)
    
    UtentiComboBox.Clear
    
    'carico i dati nella ComboBox
    For index = 2 To lastRow
        UserID = ActiveWorkbook.Sheets("Utenti").Cells(index, 1).Value
        Set user = getUtenteGeneralita(UserID)
        
        listOfUser(indexOfList, 1) = UserID
        listOfUser(indexOfList, 0) = user("Cognome") & " " & user("Nome")
        
        indexOfList = indexOfList + 1
    Next
    
    
    dummy = MultiDimensionalSortAZ(listOfUser, 0, 2, indexOfList - 1)
    UtentiComboBox.List = listOfUser
    


End Sub


' -------------------------------------------------
' Al variare della utenza, abilito i pulsanti
' -------------------------------------------------

Private Sub UtentiComboBox_Change()
    On Error GoTo eh
    
        If UtentiComboBox.Value <> "" Then

            Dim idOfUtenza As Integer
            idOfUtenza = UtentiComboBox.Column(1)
    
            'alla selezione, popolo tutti gli altri dati
            Dim Utenza As Object
            ConfermaScelta.Enabled = True
        Else
            ' se la casella torna vuota
            ConfermaScelta.Enabled = False
        End If
        
    Exit Sub
eh:
    dummy = MsgBox("Utenza non trovata o errore di battitura!" & vbCrLf & "Riprova!", vbCritical)
    
    
End Sub
