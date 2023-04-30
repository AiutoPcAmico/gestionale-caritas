Attribute VB_Name = "FunzioniGeneriche"
' ----------------------------------------------- '
' Author:           Andrea Felappi                '
' Last Modified:    03/04/2023 21.38              '
' Module:           FunzioniGeneriche             '
' ChangeLog:        Added ultimaConsegna in user  '
' Website:          aiutopcamico.altervista.org   '
' All Right Reserved                              '
' ----------------------------------------------- '


' -------------------------------------------------
' Function used while debugging
' -------------------------------------------------
Sub test()
MenuIniziale.Show
End Sub

' -------------------------------------------------
' Recupera il numero dell'ultima riga di un Sheet
' -------------------------------------------------

Public Function getLastRowIndex(workSheetName As String) As Integer
Dim righe As Integer
    righe = 0
    Do
        righe = righe + 1
    Loop Until ActiveWorkbook.Sheets(workSheetName).Cells(righe, 1).Value = ""
righe = righe - 1
getLastRowIndex = righe
End Function

' -------------------------------------------------
' Recupera la data odierna
' -------------------------------------------------

Public Function getTodayDate() As String
    Dim dtToday As String
    dtToday = Format(Date, "dd/mm/yyyy")
    getTodayDate = dtToday
End Function

' -------------------------------------------------
' Ritorna quante consegne sono state effettuate nel
' giorno di apertura attivo
' -------------------------------------------------

Public Function getNumeroConsegneOdierne() As Integer
    
    Dim lastRow As Integer
    lastRow = getLastRowIndex("Consegne")
    
    Dim DataOdierna As String
    DataOdierna = getDataOperativa
    
    Dim ConsegneOdierne As Integer
    ConsegneOdierne = 0
    
    Dim i As Integer
    For i = 2 To lastRow
        If ActiveWorkbook.Sheets("Consegne").Cells(i, 2).Value = DataOdierna Then
            ConsegneOdierne = ConsegneOdierne + 1
        End If
    Next
    
    getNumeroConsegneOdierne = ConsegneOdierne
End Function


' -------------------------------------------------
' Ritorna se il volontario loggato è Admin o User
' -------------------------------------------------

Public Function getVolontarioAuth() As String
    Dim indexOfVolontarioApertura As Integer
    indexOfVolontarioApertura = getLastRowIndex("Giornate Apertura")
    
    Dim lastVolontarioName As String
    lastVolontarioName = ActiveWorkbook.Sheets("Giornate Apertura").Cells(indexOfVolontarioApertura, 2).Value
    
    Dim indexVolontario As Integer
    indexVolontario = 1
    
    Do
        indexVolontario = indexVolontario + 1
    Loop Until ActiveWorkbook.Sheets("Volontari").Cells(indexVolontario, 1).Value = lastVolontarioName
    
    
    getVolontarioAuth = ActiveWorkbook.Sheets("Volontari").Cells(indexVolontario, 2).Value
    
    
End Function

' -------------------------------------------------
' Esegue il Sorting, data la colonna, di una intera
' listbox
' -------------------------------------------------

Public Function UserFormSortAZ(myListBox As MSForms.ListBox, colonnaDaVerificare As Integer)

    Dim cols As Long, Temp() As Variant
    cols = myListBox.ColumnCount
    ReDim Temp(0 To cols)
    With myListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.List(i, colonnaDaVerificare)) > LCase(.List(i + 1, colonnaDaVerificare)) Then
                    For c = 0 To cols - 1
                        Temp(c) = .List(i, c)
                        .List(i, c) = .List(i + 1, c)
                        .List(i + 1, c) = Temp(c)
                    Next c
                End If
            Next i
        Next j
    End With

End Function


' -------------------------------------------------
' Esegue il sorting di elementi più generici, quali
' listbox o combobox o arraymultidimensionali
' -------------------------------------------------

Public Function MultiDimensionalSortAZ(myList As Variant, colonnaDaVerificare As Integer, totalOfColonne As Integer, lunghezza As Integer)
    Dim cols As Long, Temp() As Variant
    cols = totalOfColonne
    ReDim Temp(0 To cols)
        For j = 0 To lunghezza - 1
            For i = 0 To lunghezza - 1
                If LCase(myList(i, colonnaDaVerificare)) > LCase(myList(i + 1, colonnaDaVerificare)) Then
                    For c = 0 To cols - 1
                        Temp(c) = myList(i, c)
                        myList(i, c) = myList(i + 1, c)
                        myList(i + 1, c) = Temp(c)
                    Next c
                End If
            Next i
        Next j

End Function

Public Function SortArrayAtoZ(myArray As Variant)

Dim i As Long
Dim j As Long
Dim Temp

'Sort the Array A-Z
For i = LBound(myArray) To UBound(myArray) - 1
    For j = i + 1 To UBound(myArray)
        If UCase(myArray(i)) > UCase(myArray(j)) Then
            Temp = myArray(j)
            myArray(j) = myArray(i)
            myArray(i) = Temp
        End If
    Next j
Next i

SortArrayAtoZ = myArray

End Function


' -------------------------------------------------
' Ritorna le generalità dell'utenza richiesta,
' a partire dallo Sheet "Utenti"
' -------------------------------------------------

Public Function getUtenteGeneralita(idUtenza As Integer) As Object
    totalOfUtenti = getLastRowIndex("Utenti")
    Dim utente As Object
    Set utente = New Collection
    
    rowIndexOfUtente = getUtenteRow(idUtenza)
    If rowIndexOfUtente = -1 Then
        utente.Add False, "Status"
    Else
        utente.Add True, "Status"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 1).Value, "id"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 2).Value, "Cognome"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 3).Value, "Nome"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 4).Value, "PaeseOrigine"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 5).Value, "Residenza"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 6).Value, "NumeroPersone"
        utente.Add ActiveWorkbook.Sheets("Utenti").Cells(rowIndexOfUtente, 7).Value, "NotePersonali"
    End If
    
    'recupero inoltre l'ultima consegna dell'utenza
    Dim lasrRowConsegne As Integer
    lasrRowConsegne = getLastRowIndex("Consegne")
    Dim i As Integer
    Dim dataFound As String
    
    dataFound = ""
    For i = lasrRowConsegne To 2 Step -1
        ' se la trovo, la salvo ed esco
        If ActiveWorkbook.Sheets("Consegne").Cells(i, 1).Value = utente("id") Then
            dataFound = ActiveWorkbook.Sheets("Consegne").Cells(i, 2).Value
            Exit For
        End If
    Next
    utente.Add dataFound, "UltimaConsegna"

    Set getUtenteGeneralita = utente
    
End Function


' -------------------------------------------------
' Ritorna, dato l'ID,  il numero della riga dell'
' utente nel foglio "Utenti"
' -------------------------------------------------

Public Function getUtenteRow(idUtenza) As Integer
    Set utenzaTrovata = ActiveWorkbook.Sheets("Utenti").Range("A:A").Find(What:=idUtenza)
    If Not utenzaTrovata Is Nothing Then
        getUtenteRow = utenzaTrovata.Row
    Else
        dummy = MsgBox("Attenzione!" & vbCrLf & vbCrLf & "Utente con ID " & idUtenza & " non trovato!" & vbCrLf & "Verificare la correttezza aprendo il foglio 'Utenti'" & vbCrLf & vbCrLf & "Grazie!", vbCritical)
        getUtenteRow = -1
    End If
End Function


' -------------------------------------------------
' Ritorna la data operativa (ultima apertura presente
' nel foglio "Giornate Apertura")
' -------------------------------------------------

Public Function getDataOperativa() As String
    Dim Data As String
    
    lastRowOfSheet = getLastRowIndex("Giornate Apertura")
    Data = ActiveWorkbook.Sheets("Giornate Apertura").Cells(lastRowOfSheet, 1).Value
    
    getDataOperativa = Data
End Function

' -------------------------------------------------
' Sorting alphabetically a specified Sheet
' Specialmente per le stampe
' -------------------------------------------------

Public Function SortWorksheetsTabs(sheetName As String, myRange As String, sortingColumn As String)

    ActiveWorkbook.Worksheets(sheetName).Range(myRange).Sort _
    key1:=ActiveWorkbook.Worksheets(sheetName).Range(sortingColumn), _
    Order1:=xlAscending, _
    Header:=xlNo
End Function
