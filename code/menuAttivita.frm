VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menuAttivita 
   Caption         =   "Menu Attività"
   ClientHeight    =   6975
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9015.001
   OleObjectBlob   =   "menuAttivita.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menuattivita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             Menu Attivita                 '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '

Private Sub AggiungiUtenzaButton_Click()
    Me.Hide
    ModificaUtente.Tag = -15
    ModificaUtente.Show
End Sub

Private Sub CommandButton2_Click()

End Sub


' -------------------------------------------------
' Funzione che si occupa della stampa delle consegne
' della data odierna. Copia nel foglio
' "StampaConsegneOdierne", lo esporta in PDF e lo
' salva nella cartella "stampe" della stessa directory
' -------------------------------------------------

Private Sub StampaConsegneOdierne_Click()
    Dim dataApertura As String
    Dim lastIndexOfConsegne As Integer
    Dim rowToNewSheet As Integer
    
    rowToNewSheet = 6
    lastIndexOfConsegne = getLastRowIndex("Consegne")
    dataApertura = getDataOperativa()
    
    dummy = MsgBox("Esportazione delle consegne odierne in corso!" & vbCrLf & vbCrLf & "Al termine della operazione, verrà aperto il PDF stampabile e salvato nella cartella 'stampe'", vbInformation)
    
    'svuoto il foglio per la nuova sovrascrittura
    Worksheets("StampaConsegneOdierne").Range("A6:D40").ClearContents
    
    'setto sul foglio la data odierna
    ActiveWorkbook.Sheets("StampaConsegneOdierne").Range("D3").Value = dataApertura
    
    'scorro tutta la lista delle ultime consegne e copio quelle di oggi
    Dim i As Integer
    Dim dataLetta As String
    Dim viveriConsegnati As String
    Dim oggettiConsegnati As String
    Dim idUtenza As Integer
    Dim datiUtenza As Object
    
    For i = 2 To lastIndexOfConsegne
        'verifico che la data sia uguale a quella odierna
        dataLetta = ActiveWorkbook.Sheets("Consegne").Cells(i, 2).Value
        If dataLetta = dataApertura Then
            
            'leggo e ricopio i viveri
            viveriConsegnati = ActiveWorkbook.Sheets("Consegne").Cells(i, 3).Value
            ActiveWorkbook.Sheets("StampaConsegneOdierne").Cells(rowToNewSheet, 3).Value = viveriConsegnati
            
            'leggo e ricopio i beni o vestiti
            oggettiConsegnati = ActiveWorkbook.Sheets("Consegne").Cells(i, 4).Value
            ActiveWorkbook.Sheets("StampaConsegneOdierne").Cells(rowToNewSheet, 4).Value = oggettiConsegnati
            
            'ora recupero il nome e cognome della famiglia, partendo dall'ID Utenza
            idUtenza = ActiveWorkbook.Sheets("Consegne").Cells(i, 1).Value
            'e vado a recuuperarne le proprietà
            Set datiUtenza = getUtenteGeneralita(idUtenza)
            
            'ora le salvo nel foglio
            ActiveWorkbook.Sheets("StampaConsegneOdierne").Cells(rowToNewSheet, 1).Value = datiUtenza("Cognome")
            ActiveWorkbook.Sheets("StampaConsegneOdierne").Cells(rowToNewSheet, 2).Value = datiUtenza("Nome")
            
            
            
            'incremento la prossima riga da scrivere
            rowToNewSheet = rowToNewSheet + 1
        End If
    Next
    
    'e formatto il foglio
    ActiveWorkbook.Sheets("StampaConsegneOdierne").Cells.EntireRow.AutoFit
    
    'ed infine lo stampo
    
    'modifico la data per renderla compatibile a windows
    Dim dataModificata As String
    dataModificata = Replace(dataApertura, "/", "-")
    
    ActiveWorkbook.Sheets("StampaConsegneOdierne").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ThisWorkbook.Path & "\stampe\consegne_odierne\" & dataModificata & " Consegne odierne.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    
    
End Sub

Private Sub NuovaConsegnaButton_Click()
    Me.Hide
    NuovaConsegna.Show
End Sub


Private Sub StampaConsegneUtenza_Click()
    Me.Hide
    SelezionaUtenzaStampa.Show
End Sub


' -------------------------------------------------
' Funzione che si occupa della stampa delle
' utenze . Copia nel foglio
' "StampaUtenze", lo esporta in PDF e lo
' salva nella cartella "stampe" della stessa directory
' -------------------------------------------------

Private Sub StampaUtenze_Click()
    Dim dataOdiernaEsatta As String
    Dim lastIndexOfUtenze As Integer
    Dim rowToNewSheet As Integer
    Dim i As Integer
    Dim datiUtenza As Object
    Dim idUtenza As Integer
    
    dataOdiernaEsatta = getTodayDate()
    rowToNewSheet = 6
    lastIndexOfUtenze = getLastRowIndex("Utenti")
    
    dummy = MsgBox("Esportazione delle utenze aggiornate alla data odierna in corso!" & vbCrLf & vbCrLf & "Al termine della operazione, verrà aperto il PDF stampabile e salvato nella cartella 'stampe'", vbInformation)
    
    
    'svuoto il foglio per la nuova sovrascrittura
    Worksheets("StampaUtenze").Range("A6:F40").ClearContents
    
    'setto la data di stampa del rapporto
    ActiveWorkbook.Sheets("StampaUtenze").Range("E3").Value = dataOdiernaEsatta
    
    For i = 2 To lastIndexOfUtenze
        'recupero le proprietà dell'utenza
        idUtenza = ActiveWorkbook.Sheets("Utenti").Cells(i, 1).Value
        Set datiUtenza = getUtenteGeneralita(idUtenza)
        
        'ora le salvo nel foglio
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 1).Value = datiUtenza("Cognome")
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 2).Value = datiUtenza("Nome")
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 3).Value = datiUtenza("PaeseOrigine")
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 4).Value = datiUtenza("Residenza")
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 5).Value = datiUtenza("UltimaConsegna")
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 6).Value = datiUtenza("NumeroPersone")
        ActiveWorkbook.Sheets("StampaUtenze").Cells(rowToNewSheet, 7).Value = datiUtenza("NotePersonali")
        
        rowToNewSheet = rowToNewSheet + 1
        
        'recupero l'ultima attività dell'utenza
    Next
    
    ' eseguo il sorting
    Dim rangeOfSorting As String
    rangeOfSorting = "A6:G" & rowToNewSheet - 1
    dump = SortWorksheetsTabs("StampaUtenze", rangeOfSorting, "A6")
    
    'ed infine lo stampo
    
    'modifico la data per renderla compatibile a windows
    Dim dataModificata As String
    dataModificata = Replace(dataOdiernaEsatta, "/", "-")
    
    ActiveWorkbook.Sheets("StampaUtenze").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        ThisWorkbook.Path & "\stampe\stampe_utenze\" & dataModificata & " Stampa Utenze.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    
End Sub

Private Sub TornaIndietro_Click()
    Me.Hide
    MenuIniziale.Show
End Sub

Private Sub VisualizzaConsegneOdierne_Click()
    Me.Hide
    FormConsegneOdierne.Show
End Sub

Private Sub VisualizzaUtenzeButton_Click()
    Menuattivita.Hide
    listaUtenze.Show
End Sub

' -------------------------------------------------
' Al caricamento, verifica il livello utente e,
' se necessario, mostra alcuni pulsanti aggiuntivi
' -------------------------------------------------


Private Sub UserForm_Activate()
    DataOdierna.Caption = getDataOperativa
    numConsegne.Caption = getNumeroConsegneOdierne
    
    Dim authVolontario As String
    authVolontario = getVolontarioAuth
    
    If authVolontario = "Admin" Then
        StampaUtenze.Visible = True
        StampaConsegneOdierne.Visible = True
        StampaConsegneUtenza.Visible = True
    Else
        StampaUtenze.Visible = False
        StampaConsegneOdierne.Visible = False
        StampaConsegneUtenza.Visible = False
    End If
    
End Sub

