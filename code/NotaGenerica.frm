VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NotaGenerica 
   Caption         =   "Nota"
   ClientHeight    =   3120
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6630
   OleObjectBlob   =   "NotaGenerica.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NotaGenerica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             Nota Generica                 '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '


' -------------------------------------------------
' NON VIENE USATA, MA PREFERISCO NON ELIMINARLA :)
' -------------------------------------------------

Public Function UserForms_Activate()
    Dim testo As String
    testo = Me.Tag
    
    NotaText.Caption = testo
End Function

Private Sub UserForm_Click()

End Sub
