VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NuovaConsegna 
   Caption         =   "Inserimento Nuova Consegna"
   ClientHeight    =   9360.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   18270
   OleObjectBlob   =   "NuovaConsegna.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NuovaConsegna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    03/04/2023 23.05              '
' Form:             NuovaConsegna                 '
' ChangeLog:        Added Tagliandino             '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '



Private Sub AggiungiNuovo_Click()
    Me.Hide
    ModificaUtente.Tag = -15
    ModificaUtente.Show
End Sub

Private Sub Headers_Click()

End Sub

Private Sub NumBigliettoUtenza_Change()

End Sub

' -------------------------------------------------
' Funzione che si occupa del salvataggio dei dati inseriti
' -------------------------------------------------


Private Sub SalvaConsegna_Click()
    
    Dim lastRowOfSheet As Integer
    lastRowOfSheet = getLastRowIndex("Consegne")
    
    
    'gestisco il salvataggio della consegna
    ActiveWorkbook.Sheets("Consegne").Cells(lastRowOfSheet + 1, 1).Value = UtentiComboBox.Column(1)     'userID
    ActiveWorkbook.Sheets("Consegne").Cells(lastRowOfSheet + 1, 2).Value = getDataOperativa()           'data dell'apertura odierna
    ActiveWorkbook.Sheets("Consegne").Cells(lastRowOfSheet + 1, 3).Value = ConsegnaAlimentareBox.Text   ' salvo la consegna degli alimenti
    ActiveWorkbook.Sheets("Consegne").Cells(lastRowOfSheet + 1, 4).Value = ConsegnaBeniBox.Text         'salvo la consegna dei beni
    ActiveWorkbook.Sheets("Consegne").Cells(lastRowOfSheet + 1, 5).Value = NumBigliettoUtenza.Text         'salvo il tagliando
    
    'aggiungo il tagliando giornaliero
    
    UtentiComboBox.Value = Null
    dummy = MsgBox("Consegna registrata con successo!", vbInformation, "Salvataggio")
    
End Sub




Private Sub NoteButton_Click()
    Dim idOfUtenza As Integer
    idOfUtenza = UtentiComboBox.Column(1)
    
    Dim Utenza As Object
    Set Utenza = getUtenteGeneralita(idOfUtenza)
    
    Pippo = Utenza("NotePersonali")
    MsgBox "La nota riporta: " & vbCrLf & vbCrLf & Pippo
End Sub



Private Sub TornaIndietro_Click()
    Me.Hide
    Menuattivita.Show
End Sub

' -------------------------------------------------
' Recupera i dati relativi alle ultime consegne
' dell'utente selezionato
' -------------------------------------------------

Private Sub LoadLastConsegne()
    Dim lastRowOfPage As Integer
    lastRowOfPage = getLastRowIndex("Consegne")
    
    Dim userInConsegna As Integer
    userInConsegna = UtentiComboBox.Column(1)
    
    Dim UserID As Integer
    Dim Data As String
    Dim Viveri As String
    Dim AltriBeni As String
    
    Dim listOfConsegne() As Variant
    Dim indexOfList As Integer
    indexOfList = 0
    
    
    ReDim listOfConsegne(lastRowOfPage - 1, 4)

    

    For i = lastRowOfPage To 2 Step -1
        UserID = ActiveWorkbook.Sheets("Consegne").Cells(i, 1).Value
        Data = ActiveWorkbook.Sheets("Consegne").Cells(i, 2).Value
        Viveri = ActiveWorkbook.Sheets("Consegne").Cells(i, 3).Value
        AltriBeni = ActiveWorkbook.Sheets("Consegne").Cells(i, 4).Value
        
        If UserID = userInConsegna Then
            listOfConsegne(indexOfList, 0) = UserID
            listOfConsegne(indexOfList, 1) = Data
            listOfConsegne(indexOfList, 2) = Viveri
            listOfConsegne(indexOfList, 3) = AltriBeni
        
            indexOfList = indexOfList + 1
            
        End If
        
        
    Next
    
    
    UltimeConsegneUtenza.List = listOfConsegne
    
End Sub

Private Sub UltimeConsegne_Click()

End Sub

Private Sub UltimeConsegneUtenza_Click()

End Sub

' -------------------------------------------------
' Quando si attiva la UserForm, carico gli utenti
' disponibili nella combobox
' -------------------------------------------------


Private Sub UserForm_Activate()
    UtentiComboBox.Clear
        
        
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
    
    'carico gli Headers
    Headers.RowSource = "Consegne!A1:D1"
    

    
    

End Sub

Private Sub ConsegnaAlimentareBox_Change()
    If ConsegnaAlimentareBox <> "" Or ConsegnaBeniBox <> "" Then
        SalvaConsegna.Enabled = True
    Else
        SalvaConsegna.Enabled = False
    End If
End Sub

Private Sub ConsegnaBeniBox_Change()
    If ConsegnaAlimentareBox <> "" Or ConsegnaBeniBox <> "" Then
        SalvaConsegna.Enabled = True
    Else
        SalvaConsegna.Enabled = False
    End If
End Sub


' -------------------------------------------------
' Al variare della combobox, abilito i pulsanti e
' le caselle, inoltre carico i dati richiesti
' -------------------------------------------------

Private Sub UtentiComboBox_Change()
    On Error GoTo eh
    
        If UtentiComboBox.Value <> "" Then

            Dim idOfUtenza As Integer
            idOfUtenza = UtentiComboBox.Column(1)
            
            'carico il numero della consegna giornaliera
            'verifico se devo continuare la numerazione
            Dim lastRowConsegne As Integer
            Dim lastConsegna As String
            lastRowConsegne = getLastRowIndex("Consegne")
            lastConsegna = ActiveWorkbook.Sheets("Consegne").Cells(lastRowConsegne, 2).Value
    
            If lastConsegna = getDataOperativa Then
                'se è oggi, continuo la numerazione
                NumBigliettoUtenza.Value = ActiveWorkbook.Sheets("Consegne").Cells(lastRowConsegne, 5).Value + 1
            Else
                'se è una nuova giornata
                NumBigliettoUtenza.Value = 1
            End If
    
            'alla selezione, popolo tutti gli altri dati
            Dim Utenza As Object
    
            Set Utenza = getUtenteGeneralita(idOfUtenza)
            NumeroComponenti.Caption = Utenza("NumeroPersone")
            PaeseOrigine.Caption = Utenza("PaeseOrigine")
            Residenza.Caption = Utenza("Residenza")
            LabelAlimenti = "Inserisci il ritiro di alimenti per " & vbCrLf & Utenza("Cognome") & " " & Utenza("Nome")
            LabelBeni = "Inserisci il ritiro di beni o vestiario per " & vbCrLf & Utenza("Cognome") & " " & Utenza("Nome")
            
            ConsegnaAlimentareBox.Text = ""
            ConsegnaBeniBox.Text = ""
            
            ConsegnaAlimentareBox.Enabled = True
            ConsegnaBeniBox.Enabled = True
            
            LoadLastConsegne
                       
    
            If Utenza("NotePersonali") <> "" Then
                NoteButton.Visible = True
            Else
                NoteButton.Visible = False
            End If
            
        Else
            ' se la casella torna vuota
            NumeroComponenti.Caption = ""
            PaeseOrigine.Caption = ""
            Residenza.Caption = ""
            NumBigliettoUtenza.Value = ""
            NoteButton.Visible = False
            
            LabelAlimenti = "Selezionare prima l'utenza!"
            LabelBeni = "Selezionare prima l'utenza!"
            
            ConsegnaAlimentareBox.Enabled = False
            ConsegnaBeniBox.Enabled = False
            SalvaConsegna.Enabled = False
            UltimeConsegneUtenza.Clear
        End If
        
    Exit Sub
eh:
    dummy = MsgBox("Utenza non trovata o errore di battitura!" & vbCrLf & "Riprova!", vbCritical)
    

End Sub
