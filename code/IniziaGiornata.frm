VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IniziaGiornata 
   Caption         =   "Apertura Giornaliera"
   ClientHeight    =   3360
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5880
   OleObjectBlob   =   "IniziaGiornata.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IniziaGiornata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    03/04/2023 22.20              '
' Form:             Inizio Giornata               '
' ChangeLog:        Fixed data while saving       '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '


' -------------------------------------------------
' Eseguita alla conferma, crea una nuova giornata
' nel foglio "Giornate Apertura"
' -------------------------------------------------

Private Sub ConfermaAperturaGiornata_Click()
    If VolontarioApertura.Value = "" Then
        MsgBox "Selezionare il volontario!", vbOKOnly + vbExclamation, "Attenzione!"
    Else
    
        Dim todayDate As String
                
        righe = getLastRowIndex("Giornate Apertura")
        
        ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe + 1, 1).Value = getTodayDate
        ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe + 1, 2).Value = VolontarioApertura.Value
        ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe + 1, 4).Value = "Giornata in corso"
        
        'Chiudo e torno al menu principale
        IniziaGiornata.Hide
        Menuattivita.Show
    End If
    
    'gestisco la data d'ultimo accesso
    'trovo la riga dell'utente che sta facendo l'accesso
    Dim i As Integer
    Dim lastVolontarioRow As Integer
    Dim volontarioLetto As String
    
    lastVolontarioRow = getLastRowIndex("Volontari")
    For i = 2 To lastVolontarioRow
        volontarioLetto = ActiveWorkbook.Sheets("Volontari").Cells(i, 1).Value
        If VolontarioApertura.Value = volontarioLetto Then
            ActiveWorkbook.Sheets("Volontari").Cells(i, 3).Value = getTodayDate
        End If
    Next

    
End Sub


Private Sub DataOdierna_Click()

End Sub

Private Sub TornaIndietro_Click()
    Me.Hide
    MenuIniziale.Show
End Sub

' -------------------------------------------------
' Al loading, carica i volontari selezionabili
' -------------------------------------------------

Private Sub UserForm_Activate()

    VolontarioApertura.Clear
    DataOdierna.Caption = getTodayDate
    
    Dim righe As Integer
    righe = getLastRowIndex("Volontari")
   
    Dim i As Integer
    For i = 2 To righe
        Dim nome As String
    
        nome = ActiveWorkbook.Sheets("Volontari").Cells(i, 1).Value
    
        With VolontarioApertura
            .AddItem (nome)
        End With
    Next
 
End Sub


