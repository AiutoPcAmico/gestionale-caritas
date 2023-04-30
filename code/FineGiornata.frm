VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FineGiornata 
   Caption         =   "Chiusura Giornaliera"
   ClientHeight    =   3615
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5625
   OleObjectBlob   =   "FineGiornata.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FineGiornata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             Fine Giornata                 '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '

' -------------------------------------------------
' Eseguita alla pressione della conferma, salva
' il volontario e lo stato di chiusura nel foglio
' "Giornate Apertura"
' -------------------------------------------------


Private Sub ConfermaFineGiornata_Click()
    

    If VolontarioChiusura.Value = "" Then
        MsgBox "Selezionare il volontario!", vbOKOnly + vbExclamation, "Attenzione!"
    Else
        righe = getLastRowIndex("Giornate Apertura")
        ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe, 3).Value = VolontarioChiusura.Value
        ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe, 4).Value = "Giornata terminata correttamente"
        
        'Chiudo e torno al menu principale
        FineGiornata.Hide
        MenuIniziale.Show
    End If
    
    

    
End Sub


Private Sub TornaIndietro_Click()
    Me.Hide
    MenuIniziale.Show
End Sub


' -------------------------------------------------
' Al load della UserForm, carica i volontari disponibili
' e la data odierna nella Label
' -------------------------------------------------

Private Sub UserForm_Activate()

    VolontarioChiusura.Clear
    DataOdierna.Caption = getDataOperativa
    
    Dim righe As Integer
    righe = getLastRowIndex("Volontari")
   
    Dim i As Integer
    For i = 2 To righe
        Dim nome As String
    
        nome = ActiveWorkbook.Sheets("Volontari").Cells(i, 1).Value
    
        With VolontarioChiusura
            .AddItem (nome)
        End With
    Next
 
End Sub

Private Sub VolontarioChiusura_Change()

End Sub
