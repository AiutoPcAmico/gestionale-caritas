VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuIniziale 
   Caption         =   "Menu Iniziale"
   ClientHeight    =   4965
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7200
   OleObjectBlob   =   "MenuIniziale.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MenuIniziale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             MenuIniziale                  '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '


Private Sub ContinuaGiornata_Click()
    Me.Hide
    Menuattivita.Show
End Sub

Private Sub GoToApriGiornata_Click()
MenuIniziale.Hide
IniziaGiornata.Show
End Sub


Private Sub GoToTerminaGiornata_Click()

MenuIniziale.Hide
FineGiornata.Show
End Sub


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

' -------------------------------------------------
' Al loading, carica i dati richiesti per iempire
' le lables presenti. Varia inoltre i pulsanti disponibili
' -------------------------------------------------

Private Sub UserForm_Activate()
    'recupero i dati dell'ultima chiusura
    righe = getLastRowIndex("Giornate Apertura")
    Dim lastDay As String
    Dim lastClose As String
    Dim lastVolontario As String
    lastDay = getDataOperativa
    lastClose = ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe, 4).Value
    lastVolontario = ActiveWorkbook.Sheets("Giornate Apertura").Cells(righe, 3).Value
    
    
    ' definisco, dall'ultimo stato la situazione del menu
    Select Case lastClose
        Case "Giornata in corso"
            ContinuaGiornata.Visible = True
            GoToApriGiornata.Visible = False
            GoToTerminaGiornata.Visible = True
            StatoChiusuraLabel.BackColor = RGB(255, 192, 192)
            
        Case "Giornata terminata correttamente"
            ContinuaGiornata.Visible = False
            GoToApriGiornata.Visible = True
            GoToTerminaGiornata.Visible = False
            StatoChiusuraLabel.BackColor = RGB(145, 255, 145)
            
        Case Else
            ' non dovrebbe mai capitare, ma si sa mai...
            ContinuaGiornata.Visible = False
            GoToApriGiornata.Visible = True
            GoToTerminaGiornata.Visible = True
            StatoChiusuraLabel.BackColor = RGB(145, 145, 255)
            
    End Select
    
    
    
    UltimaAperturaLabel.Caption = lastDay
    ContinuaGiornata.Caption = "Continua la giornata" & vbCrLf & "del " & lastDay
    StatoChiusuraLabel.Caption = lastClose
    VolontarioChiusuraLabel = lastVolontario
    
End Sub
