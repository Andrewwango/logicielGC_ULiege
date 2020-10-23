Attribute VB_Name = "tableManagement"
Option Explicit
'   LOGICIELGC 2018 tableManagement   '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '
Public selectedRow As String
Public selectedIndex As Integer

Public Const DATEFORMAT As String = "DD/MM/YYYY"

'THESE NEED TO BE KEPT UP TO DATE
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Const COLUMNCLIENTID As Integer = 1
Public Const COLUMNCLIENTNOM As Integer = 2
Public Const COLUMNCLIENTADRESSE As Integer = 3
Public Const COLUMNCLIENTREMARQUES As Integer = 4

Public Const COLUMNESSAIID As Integer = 1
Public Const COLUMNESSAITYPE As Integer = 2
Public Const COLUMNESSAIVERSION As Integer = 3
Public Const COLUMNESSAISORTILEDATE As Integer = 4
Public Const COLUMNESSAIQUANTITY As Integer = 11
Public Const COLUMNESSAISDEMANDES As Integer = 15
Public Const COLUMNESSAINORME As Integer = 16
Public Const COLUMNESSAIREMARQUES As Integer = 17
'''''''''''''''''''''''''''''''''''''''''''''''''

Function cellInTable(mycell As Range) As Boolean
    cellInTable = False
    On Error Resume Next
    cellInTable = (mycell.ListObject.Name <> "")
    On Error GoTo 0
End Function
Public Function findInColumn(searchvalue As Variant, columnno As Integer, searchsheet As String, _
    searchTable As String, Optional direction As XlSearchDirection = xlPrevious) As Range
'Search for a value based on one column
'Can be used like findInColumn(etc).EntireRow to obtain associated info
'Returns Nothing type if nothing found
Set findInColumn = ActiveWorkbook.Worksheets(searchsheet).ListObjects(searchTable).DataBodyRange.Columns(columnno).Find(searchvalue, searchDirection:=direction)

End Function

Public Function concatenateRow(irow As Range, start As Integer, Optional connector As String = vbCrLf) As String
'Return string of cell values of irow concatenated together, starting at column start
    Dim i As Integer: i = start
    Do
        concatenateRow = concatenateRow & irow.Cells(1, i) & connector
        i = i + 1
    Loop Until i = irow.CurrentRegion.Columns.count
End Function

Public Sub searchTable(searchsheet As String, searchTable As String, textboxes As Collection, Optional wb As Workbook = Nothing)
'Finds first content in textboxes on form, searches table for content,
'then populates other textboxes with row data
    
    Dim foundSomething As Boolean: foundSomething = False
    Dim secondPass As Boolean: secondPass = False
    Dim resultsCollection As New Collection 'results are strings of row.address
    Dim resultsCollectionNarrow As New Collection
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    With wb.Worksheets(searchsheet)
        Dim table As ListObject
        Set table = .ListObjects(searchTable)
        
        'Iterate through textboxes
        Dim iter As Integer
        Dim itxt As Variant
        For iter = 1 To textboxes.count
            Set itxt = textboxes(iter)
            If itxt.Value <> "" Then
                
                'Search value in table
                'Search backwards to get latest versions first
                With table.DataBodyRange.Columns(iter)
                    
                    Dim searchresult As Range
                    Set searchresult = .Find(itxt.Value, searchDirection:=xlPrevious)
                    
                    If Not searchresult Is Nothing Then
                        
                        'Found first result
                        foundSomething = True
                        
                        'firstrow so when find loops back to beginning, it knows to stop
                        Dim firstrow As Range
                        Set firstrow = searchresult.EntireRow

                        'Populate collection of results
                        Do
                            'Only add if essaiID not old version ("!xxx") This only applied to searching in essaitable!
                            If Not searchTable = "essaisTable" Or InStr(searchresult.EntireRow.Cells(1, COLUMNESSAIID).Value, "!") = 0 Then
                                
                            If secondPass = False Then
                                'Add result as string to collection
                                resultsCollection.Add searchresult.EntireRow.Address
                                
                            Else
                                'Narrow down if it's on second pass
                                Dim iterresult As Variant
                                For Each iterresult In resultsCollection
                                    If searchresult.EntireRow.Address = CStr(iterresult) Then _
                                        resultsCollectionNarrow.Add searchresult.EntireRow.Address
                                    'This way, the final collection is the intersection of the results from two searches
                                Next iterresult
                            End If
                            End If
                            
                            'Search again for next row
                            Set searchresult = .FindPrevious(searchresult)
                            If searchresult Is Nothing Then Exit Do
                            
                        Loop While searchresult.EntireRow.Address <> firstrow.Address
                    Else
                        'Check next box
                        foundSomething = False
                    End If
                End With
                
                'If too many results then loop again with next box and narrow results down
                If secondPass = False And resultsCollection.count > 10 Then
                    secondPass = True
                    foundSomething = False
                End If
                
                'Finished searching if something was found
                If foundSomething = True Then
                    Exit For
                End If
                      
            End If
            
            'If no search results, check for more textboxes with content in
        Next iter
    End With
    
    
    'Done searching, time to display results!
    
    If secondPass = True Then
        foundSomething = True
        If resultsCollectionNarrow.count > 0 Then Set resultsCollection = resultsCollectionNarrow
    End If
       
    If foundSomething = True Then
    
        Load frmMultipleResults
        With frmMultipleResults.lbxMultipleResults
            .ColumnCount = 2
            .ColumnWidths = "1;1000"
        End With
        
        'Add all results from collection into listbox
        Dim resultitem As Variant
        For Each resultitem In resultsCollection
        
            With frmMultipleResults.lbxMultipleResults
                .AddItem (resultitem)
                .List(.ListCount - 1, 1) = concatenateRow(wb.Worksheets(searchsheet).Range(resultitem), 1, connector:=" | ")
            End With
        Next resultitem
        
        'Allow user to choose one
        frmMultipleResults.Show
                              
        'If nothing selected just quit
        If selectedRow = "" Then Exit Sub
                              
        'Populate other textboxes with selected row
        With wb.Worksheets(searchsheet).Range(selectedRow)
            Dim j As Integer
            For j = 1 To textboxes.count
                textboxes(j).Text = .Cells(1, j)
            Next j
        End With
        
        'Done
        selectedRow = ""
        Exit Sub
    
    Else
        MsgBox "Rien trouvé!!", Buttons:=vbExclamation, Title:="Rechercher"
    End If
        
End Sub

Public Sub clearAll(textboxes As Collection)
    Dim itxt As Variant
    For Each itxt In textboxes
        itxt.Text = ""
        itxt.Locked = False
    Next itxt
End Sub

Public Sub saveToTable(savesheet As String, savetable As String, textboxes As Collection, _
    Optional modifyrow As Range = Nothing, Optional newRowAfter As Range = Nothing)
'Save textbox details as a row in table,
'The order of the table headings must be the same as the corresponding textboxes in the collection
    
    Dim table As ListObject
    Set table = Worksheets(savesheet).ListObjects(savetable)
    Dim myRow As Range
    
    'New row in table or modify existing one?
    If modifyrow Is Nothing And newRowAfter Is Nothing Then
        Set myRow = table.ListRows.Add.Range
    ElseIf newRowAfter Is Nothing Then
        Set myRow = modifyrow
    Else
        Set myRow = table.ListRows.Add(newRowAfter.Row).Range
    End If

    Dim i As Integer
    For i = 1 To textboxes.count
        myRow.Cells(1, i) = textboxes(i).Text
    Next i
    
    MsgBox "Enregistré!", Buttons:=vbInformation
End Sub

Public Function latestEssaiID() As Long
'Return last row's essaiID - the essaiID column should be in ascending order!!
    Dim myTable As ListObject
    Set myTable = Worksheets("Essais").ListObjects("essaisTable")

    latestEssaiID = myTable.ListRows(myTable.ListRows.count).Range.Cells(1, COLUMNESSAIID).Value
End Function

Public Function latestClientID(firstLetter As String) As String
'Returns last ClientID starting with the desired letter

    Dim myTable As ListObject
    Set myTable = Worksheets("Clients").ListObjects("clientsTable")
    Dim filteredRange As Range
    
    'Filter to get IDs starting with desired letter
    myTable.Range.AutoFilter Field:=COLUMNCLIENTID, Criteria1:="=" & firstLetter & "*"
    
    'Get filter results (if there are any)
    On Error Resume Next
    Set filteredRange = myTable.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If filteredRange Is Nothing Then
        latestClientID = "00000"
    ElseIf filteredRange.Rows.count > 1 Then
        'get last one
        latestClientID = filteredRange.End(xlDown).Cells(1, COLUMNCLIENTID).Value
    ElseIf filteredRange.Rows.count = 1 Then
        latestClientID = filteredRange.Rows.Cells(1, COLUMNCLIENTID).Value
    End If
    
    myTable.Range.AutoFilter Field:=COLUMNCLIENTID 'Clear filter
End Function

Public Sub formatClientID()
'Goes through a Table and renames all the IDs e.g. from A100 -> A0100

    Dim clientIDFormatString As String: clientIDFormatString = "0000"
    Dim coln As Integer: coln = 1
    
    Dim firstChar As String
    Dim myID As String
    With Worksheets("Clients").ListObjects("clientsTable").DataBodyRange
        Dim i As Integer
        For i = 1 To .Rows.count
            myID = .Cells(i, coln).Value
            firstChar = UCase(Left(myID, 1))
            
            On Error Resume Next
            .Cells(i, coln).Value = firstChar & Format$(Right(myID, Len(myID) - 1), clientIDFormatString)
        Next i
    End With
End Sub

Public Sub stripField()
'Goes through a Table and strips some characters from start every cell in column e.g. 4ENM -> ENM

    Dim noOfChars As Integer: noOfChars = 1
    Dim coln As Integer: coln = 2
    
    With Worksheets("Essais").ListObjects("essaisTable").DataBodyRange
        Dim i As Integer
        For i = 1 To .Rows.count
                       
            On Error Resume Next
            .Cells(i, coln).Value = Right(.Cells(i, coln).Value, Len(.Cells(i, coln).Value) - noOfChars)
        Next i
    End With
End Sub

