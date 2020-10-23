Attribute VB_Name = "wordTools"
Option Explicit
'   LOGICIELGC 2018 wordTools    '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '

Public Sub wRT(Document As Variant, tfind As String, treplace As String, Optional moveSelection As Boolean = False)
'wordReplaceText
'Used for replacing <<>> fields in PV document templates
    Dim r As Variant: Set r = Document.Range
    With r
        .Find.Execute FindText:=tfind, ReplaceWith:=treplace
        If moveSelection = True And .Find.Found Then .Select
    End With
End Sub

Public Sub eRT(mysheet As Variant, tfind As String, treplace As String)
'excelReplaceText
    mysheet.Cells.Replace what:=tfind, Replacement:=treplace, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
End Sub

Public Sub wordRemovePagesFromEnd(wordapp As Variant, Document As Variant, numberOfPages As Integer)
    Dim rgePages As Variant
    Dim PageCount As Integer
    PageCount = Document.ComputeStatistics(2)
    
    With wordapp.Selection
    
        .Goto what:=1, Which:=1, count:=PageCount - numberOfPages + 1
        Set rgePages = .Range
        
        .Goto what:=1, Which:=1, count:=PageCount
        
        rgePages.End = .Bookmarks("\Page").Range.End
        rgePages.Delete
        .TypeBackspace
    End With
End Sub
