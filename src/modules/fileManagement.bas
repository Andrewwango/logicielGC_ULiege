Attribute VB_Name = "fileManagement"
Option Explicit
'   LOGICIELGC 2018 fileManagement   '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '

Public DOCUMENTSROOTPATH As String
Public FICHEVERTEFOLDER As String: Public PAGEDEGARDEFOLDER As String: Public FICHERESULTATSFOLDER As String
Public TEMPLATESFOLDER As String: Public LEGACYWORKBOOKSFOLDER As String
Public FICHEVERTETEMPLATE As String: Public PAGEDEGARDETEMPLATE As String
Public FICHEVERTENAMEFORMAT As String: Public PAGEDEGARDENAMEFORMAT As String: Public FICHERESULTATSNAMEFORMAT As String
Public LEGACYESSAISFILENAME As String: Public LEGACYCLIENTSFILENAME As String
Public REtemplateArray As Variant

Public Sub initialiseFilePaths()
    DOCUMENTSROOTPATH = ThisWorkbook.path & "\Documents"
    FICHEVERTEFOLDER = DOCUMENTSROOTPATH
    PAGEDEGARDEFOLDER = DOCUMENTSROOTPATH
    FICHERESULTATSFOLDER = DOCUMENTSROOTPATH
    
    TEMPLATESFOLDER = ThisWorkbook.path & "\Templates"
    LEGACYWORKBOOKSFOLDER = ThisWorkbook.path & "\Legacy"
    
    FICHEVERTETEMPLATE = "TemplateFicheVerte.dot"
    PAGEDEGARDETEMPLATE = "TemplatePageDeGarde.dot"
    'FicheResultats Template names are managed below
    
    FICHEVERTENAMEFORMAT = "FicheVerte"
    PAGEDEGARDENAMEFORMAT = "PageDeGarde"
    FICHERESULTATSNAMEFORMAT = "FicheResultats"
    
    LEGACYESSAISFILENAME = "LegacyEssais250918.xls"
    LEGACYCLIENTSFILENAME = "LegacyClients250918.xls"

    'FORMAT:
    'Array(  templatefichier,  type,  matériel,  essaispécifique,  1er ligne de table de résultats  ), _
    '
    
    REtemplateArray = Array( _
 _
 _
    Array("TemplateResultatsDefault.xlt", "Default", "", "", "14"), _
    Array("TemplateResultatsENMCubesCompression.xlt", "ENM", "Cubes", "Compression", "14"), _
    Array("TemplateResultatsDefault.xlt", "ENM", "Cubes", "Absorption", "14"), _
    Array("TemplateResultatsDefault.xlt", "ENM", "Carottes", "Compression", "14"), _
    Array("TemplateResultatsDefault.xlt", "ENM", "Carottes", "Absorption", "14"), _
    Array("TemplateResultatsDefault.xlt", "EM", "", "Traction", "14"), _
    Array("TemplateResultatsDefault.xlt", "ENM", "Cylindres", "Compression", "14") _
 _
 _
    )


End Sub

Function CreateDocumentFileName(doctype As String, docid As String, Optional docversion As Integer = 0, _
    Optional docext As String = ".doc") As String
'Create file name with v number if more than 1 and nothing if ficheverte or resultatsessai
    Dim rootpath As String
    If doctype = "RE" Then
        CreateDocumentFileName = docid & " " & FICHERESULTATSNAMEFORMAT
        rootpath = FICHERESULTATSFOLDER
    ElseIf doctype = "FV" Then
        CreateDocumentFileName = docid & " " & FICHEVERTENAMEFORMAT
        rootpath = FICHEVERTEFOLDER
    ElseIf docversion > 1 Then
        CreateDocumentFileName = docid & " " & PAGEDEGARDENAMEFORMAT & " v" & docversion
        rootpath = PAGEDEGARDEFOLDER
    Else
        CreateDocumentFileName = docid & " " & PAGEDEGARDENAMEFORMAT
        rootpath = PAGEDEGARDEFOLDER
    End If
    
    CreateDocumentFileName = rootpath & "\" & CreateDocumentFileName & docext
End Function

