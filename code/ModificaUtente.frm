VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModificaUtente 
   Caption         =   "Modifica un Utente"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "ModificaUtente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModificaUtente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             Modifica Utente               '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '


' -------------------------------------------------
' Chiamata da Activate  o da AnnullaModifiche,
' legge i dati dell'utente e compila i campi
' -------------------------------------------------

Private Function loadValues(idOfUser As Integer)
    'resetto i colori
    NomeBox.BackColor = RGB(255, 255, 255)
    CognomeBox.BackColor = RGB(255, 255, 255)
    ResidenzaBox.BackColor = RGB(255, 255, 255)
    PaeseOrigineComboBox.BackColor = RGB(255, 255, 255)

            
    Dim UtenteDaModificare As Object
    Set utente = getUtenteGeneralita(idOfUser)
    If utente.Item("Status") Then
        NomeBox.Text = utente.Item("Nome")
        CognomeBox.Text = utente.Item("Cognome")
        PaeseOrigineComboBox.Value = utente.Item("PaeseOrigine")
        ResidenzaBox.Text = utente.Item("Residenza")
        NumeroPersoneBox.Text = utente.Item("NumeroPersone")
        EventualiNoteBox.Text = utente.Item("NotePersonali")
        
    Else
        Me.Hide
        Menuattivita.Show
    End If
    
    
    
        
End Function



Private Sub AnnullaModifiche_Click()
    Dim idOfUser As Integer
    idOfUser = Me.Tag
    If idOfUser = -15 Then
            ' sto aggiungendo un nuovo utente
        Else
            ' modifico uno già esistente
            loadValues (idOfUser)
        End If
End Sub

' -------------------------------------------------
' Funzione che si occupa del salvataggio dei dati
' inseriti
' -------------------------------------------------

Private Function ModifyExistingUtente(idUtente As Integer)
' procedo al salvataggio
        Dim rowNumber As Integer
        rowNumber = getUtenteRow(idUtente)
        
        If rowNumber <> -1 Then
            'se è stato trovato (cioè, in teoria, sempre)
            ActiveWorkbook.Sheets("Utenti").Cells(rowNumber, 2).Value = CognomeBox.Text
            ActiveWorkbook.Sheets("Utenti").Cells(rowNumber, 3).Value = NomeBox.Text
            ActiveWorkbook.Sheets("Utenti").Cells(rowNumber, 4).Value = PaeseOrigineComboBox.Value
            ActiveWorkbook.Sheets("Utenti").Cells(rowNumber, 5).Value = ResidenzaBox.Text
            ActiveWorkbook.Sheets("Utenti").Cells(rowNumber, 6).Value = NumeroPersoneBox.Text
            ActiveWorkbook.Sheets("Utenti").Cells(rowNumber, 7).Value = EventualiNoteBox.Text
            
        Else
            'nulla: stampo già nella funzione
        End If
End Function

' -------------------------------------------------
' Si occupa del salvataggio nel caso in cui sia
' un nuovo utente da aggiungere
' -------------------------------------------------

Private Function AddNewUtente()
    lastRow = getLastRowIndex("Utenti")
    
    Dim lastID As Integer
    lastID = ActiveWorkbook.Sheets("Utenti").Cells(lastRow, 1).Value

    'salvo i dati
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 1).Value = lastID + 1
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 2).Value = CognomeBox.Text
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 3).Value = NomeBox.Text
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 4).Value = PaeseOrigineComboBox.Text
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 5).Value = ResidenzaBox.Text
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 6).Value = NumeroPersoneBox.Text
    ActiveWorkbook.Sheets("Utenti").Cells(lastRow + 1, 7).Value = EventualiNoteBox.Text
    
    Me.Tag = lastID + 1
End Function

' -------------------------------------------------
' Alla conferma, verifica che i campi obbligatori
' siano compilati e chiama la funzione in base se
' viene aggiunto un nuovo utente oppure modificato
' -------------------------------------------------

Private Sub ConfermaButton_Click()
    'resetto i colori
    NomeBox.BackColor = RGB(255, 255, 255)
    CognomeBox.BackColor = RGB(255, 255, 255)
    ResidenzaBox.BackColor = RGB(255, 255, 255)
    PaeseOrigineComboBox.BackColor = RGB(255, 255, 255)
    
    Dim idOfUser As Integer
    idOfUser = Me.Tag

    'effettuo il controllo che tutti i campi obbligatori siano presenti
    Dim canISave As Boolean
    canISave = True
        
    If NomeBox.Text = "" Then
        canISave = False
        NomeBox.BackColor = RGB(255, 192, 192)
    End If
    
    If CognomeBox.Text = "" Then
        canISave = False
        CognomeBox.BackColor = RGB(255, 192, 192)
    End If
    
    If NomeBox.Text = "" Then
        canISave = False
        NomeBox.BackColor = RGB(255, 192, 192)
    End If
    
    If ResidenzaBox.Text = "" Then
        canISave = False
        ResidenzaBox.BackColor = RGB(255, 192, 192)
    End If
    
    If PaeseOrigineComboBox.Text = "" Then
        canISave = False
        PaeseOrigineComboBox.BackColor = RGB(255, 192, 192)
    End If
    
    'dopo tutti i controlli
    If canISave Then
        If idOfUser = -15 Then
            ' se è una nuova aggiunta
            AddNewUtente
        Else
            ' se si tratta di una modifica di un utente già esistente
            ModifyExistingUtente (idOfUser)
        End If
        
    dummy = MsgBox("Salvataggio avvenuto correttamente!" & vbCrLf & "Verrai ora reindirizzato alla lista delle utenze", vbInformation, "Salvato")
    Me.Hide
    listaUtenze.Show
        
    Else
        ' mando un messaggio di errore
        dummy = MsgBox("Attenzione!" & vbCrLf & vbCrLf & "Non sono stati compilati tutti i campi obbligatori!" & vbCrLf & "Prego, verificare", vbExclamation, "Errore!")
    End If
    
    
    
    
End Sub



Private Sub PaeseOrigineComboBox_Change()

End Sub

Private Sub TornaIndietro_Click()
    Me.Hide
    Menuattivita.Show
End Sub



' -------------------------------------------------
' FAl caricamento, verifica se si tratta di un nuovo
' inserimento o di una modifica, e compila eventualmente
' i campi.
' -------------------------------------------------

Private Sub UserForm_Activate()
    'svuoto tutti i campi, altrimenti nel crearne uno nuovo dopo la modifica,
    ' se li risalva
    
    NomeBox.Text = ""
    CognomeBox.Text = ""
    PaeseOrigineComboBox.Clear
    ResidenzaBox.Text = ""
    NumeroPersoneBox.Text = ""
    EventualiNoteBox.Text = ""
    
    Dim idOfUser As Integer
    
    Dim stati As Variant
    stati = Array("Afghanistan", "Bolivia", "Brasile", "Camerun", "Ecuador", "Gambia", "Ghana", "India", "Italia", "Libia", "Mali", "Marocco", "Nigeria", "Pakistan", "Perù", "Senegal", "Sierra Leone", "Somalia", "Tunisia", "Ucraina", "Bulgaria", "Romania", "Burkina Faso", "Benin", "Niger", "Egitto", "Congo", "Kenya", "Iraq", "Iran", "Turchia", "Russia")
    
    'Inserisco la lista degli stati d'origine
    stati = SortArrayAtoZ(stati)
    
    Dim i As Integer
    For i = 0 To UBound(stati) - LBound(stati)
        PaeseOrigineComboBox.AddItem stati(i)
    Next
    
    
    
    
    
    
    
    If Me.Tag <> "" Then
        idOfUser = Me.Tag
        
        If idOfUser = -15 Then
            ' sto aggiungendo un nuovo utente
            'non faccio nulla
            
        Else
            ' modifico uno già esistente
            loadValues (idOfUser)
        End If
        
    Else
        dummy = MsgBox("Questo modulo può essere aperto soltanto da un altro menu." & vbCrLf & vbCrLf & "Prego, riaprire il menu iniziale!", vbCritical)
       Me.Hide
       Menuattivita.Show
    End If
        
End Sub
