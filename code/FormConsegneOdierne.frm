VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormConsegneOdierne 
   Caption         =   "Visualizza Consegne Odierne"
   ClientHeight    =   6165
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   17370
   OleObjectBlob   =   "FormConsegneOdierne.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormConsegneOdierne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    03/04/2023 23.05              '
' Form:             FormConsegneOdierne           '
' ChangeLog:        Added Tagliandino             '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '

Private Sub ConsegneOdierne_Click()

End Sub

Private Sub TornaIndietro_Click()
    Me.Hide
    Menuattivita.Show
End Sub


' -------------------------------------------------
' Al load della Form, carica tutte le consegne
' avvenute nel giorno di apertura ultimo.
' Esegue inoltre il sorting in ordine alfabetico
' -------------------------------------------------

Private Sub UserForm_Activate()

    'carico label con numero di consegne odierne
    TotConsegne.Caption = getNumeroConsegneOdierne
    
'carico gli Headers
    Headers.AddItem
    Headers.List(0, 0) = "Num"
    Headers.List(0, 1) = "Utenza"
    Headers.List(0, 2) = "Consegna Alimentare"
    Headers.List(0, 3) = "Consegna di Beni e Vestiario"
    
    'carico le consegne odierne
    Dim lastRowOfPage As Integer
    lastRowOfPage = getLastRowIndex("Consegne")
    
    Dim UserID As Integer
    Dim Viveri As String
    Dim AltriBeni As String
    Dim datiUtente As Object
    Dim numTagliandino As Integer
    
    Dim listOfConsegne() As Variant
    Dim indexOfList As Integer
    indexOfList = 0
    
    Dim DataOperativa As String
    DataOperativa = getDataOperativa
    
    ' ridefinisco la grandezza per garantire il caricamento dei dati
    ReDim listOfConsegne(lastRowOfPage - 1, 4)

    ' leggo ogni Record
    For i = lastRowOfPage To 2 Step -1
        UserID = ActiveWorkbook.Sheets("Consegne").Cells(i, 1).Value
        Data = ActiveWorkbook.Sheets("Consegne").Cells(i, 2).Value
        Viveri = ActiveWorkbook.Sheets("Consegne").Cells(i, 3).Value
        AltriBeni = ActiveWorkbook.Sheets("Consegne").Cells(i, 4).Value
        numTagliandino = ActiveWorkbook.Sheets("Consegne").Cells(i, 5).Value
        
        ' se è odierno, lo inserisco nella ListBox
        If Data = DataOperativa Then
            
            Set datiUtente = getUtenteGeneralita(UserID)
            listOfConsegne(indexOfList, 0) = numTagliandino
            listOfConsegne(indexOfList, 1) = datiUtente("Cognome") & " " & datiUtente("Nome")
            listOfConsegne(indexOfList, 2) = Viveri
            listOfConsegne(indexOfList, 3) = AltriBeni
            
        
            indexOfList = indexOfList + 1
            
        End If
        
        
    Next
    ConsegneOdierne.List = listOfConsegne
End Sub
