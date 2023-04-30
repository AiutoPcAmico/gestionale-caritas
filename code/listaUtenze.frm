VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} listaUtenze 
   Caption         =   "Visualizza le Utenze"
   ClientHeight    =   7305
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10980
   OleObjectBlob   =   "listaUtenze.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "listaUtenze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    29/03/2023 16.05              '
' Form:             listaUtenze                   '
' ChangeLog:        First Release                 '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '

' -------------------------------------------------
' Funzione per eliminare l'utenza selezionata
' -------------------------------------------------

Private Sub EliminaUtenza_Click()
    ' nella realtà di cui ho preso esempio, non è possibile questa operazione
    ' in quanto comunque tutte le utenze vengono definite come "storiche" ma MAI eliminate.
    ' Inoltre, per scelta, NON CANCELLO eventuali consegne dell'utente, in quanto è utile
    ' per tenere traccia del magazzino alimentare e dei beni materiali.
    
    Debug.Print ("Pippo")
    idToDelete = Me.ListaUtenti
    Debug.Print (TypeName(idToDelete))
    Dim utente As Object
    Set utente = getUtenteGeneralita(CInt(idToDelete))
    
    If utente.Item("Status") Then
       
        Response = MsgBox("Sei proprio sicuro di voler eliminare l'utenza " & utente.Item("Cognome") & " " & utente.Item("Nome"), vbQuestion + vbYesNo, "Elimina Utenza")
        Debug.Print (Response)
    
        If Response = vbYes Then
            MsgBox "Nell'applicativo che andrà in produzione, questa funzione non è implementata, in quanto comunque tutte le utenze vengono definite come storiche ma MAI eliminate." & vbCrLf & vbCrLf & "Inoltre, per scelta, NON CANCELLO eventuali consegne dell'utente, in quanto è utile per tenere traccia del magazzino e dei viveri." & vbCrLf & vbCrLf & "Solo per scopo d'esercizio, eliminerò la riga dell'utente selezionato.", vbCritical
            ActiveWorkbook.Sheets("Utenti").Rows(getUtenteRow(idToDelete)).EntireRow.Delete
            
            listaUtenze.Hide
            listaUtenze.Show
        End If
    End If
End Sub


' -------------------------------------------------
' Dopo la selezione di un elemento della lista,
' abilito i pulsanti
' -------------------------------------------------

Private Sub ListaUtenti_AfterUpdate()
  strState = Me.ListaUtenti
  If strState <> "" Then
    EliminaUtenza.Enabled = True
    modificaAnagraficaButton.Enabled = True
  Else
    EliminaUtenza.Enabled = False
    modificaAnagraficaButton.Enabled = False
  End If
  
  ' ListaUtenti.Selected(ListaUtenti.ListIndex) = False
  ' MsgBox ("E ' stato selezionato l'utente" & Str(strState))
  
End Sub



Private Sub modificaAnagraficaButton_Click()
    idToModify = Me.ListaUtenti
    Me.Hide
    ModificaUtente.Tag = idToModify
    ModificaUtente.Show
    
End Sub



Private Sub TornaIndietro_Click()
    Me.Hide
    Menuattivita.Show
End Sub


' -------------------------------------------------
' Al caricamento della UserForm, carico tutti i dati
' per l'header e per le utenze presenti nel gestionale
' -------------------------------------------------

Private Sub UserForm_Activate()


    ' Ho dovuto fare 2 listBox differenti
    ' Non è il massimo, ma se voglio utilizzare le funzioni di sorting, necessito di "copiare" i dati nella ListBox
    ' e non solamente di effettuare il binding con RowSource. (necessità di manipoliazione dei dati)
    ' Utilizzando la funzione List, copio i dati correttamente nella ListBox.
    ' Purtroppo non permette di specificare gli Headers....
    
    Dim columnDimensions As String
    columnDimensions = "0;100;100;70;100;40;10"
    
    Headers.ColumnCount = 6
    Headers.ColumnWidths = columnDimensions
    Headers.RowSource = "Utenti!A1:F1"
    
    
    modificaAnagraficaButton.Enabled = False
    EliminaUtenza.Enabled = False


    Dim riga As Integer
    riga = getLastRowIndex("Utenti")
    
    ListaUtenti.ColumnCount = 6
    
    ListaUtenti.ColumnWidths = columnDimensions
    
    myRange = "A2:F" & LTrim(Str(riga))
    ListaUtenti.List = Sheets("Utenti").Range(myRange).Value
    
    
    dummy = UserFormSortAZ(ListaUtenti, 1)
        
    'verifico se l'utente loggato può eliminare utenze
    typeAuth = getVolontarioAuth
    If typeAuth = "Admin" Then
        EliminaUtenza.Visible = True
    Else
        EliminaUtenza.Visible = False
    End If
    
End Sub

