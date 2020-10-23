VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEssais 
   Caption         =   "Essais"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   OleObjectBlob   =   "frmEssais.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEssais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   LOGICIELGC 2018 frmEssais    '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '

Dim myTextboxes As New Collection

Private Sub btnActions_Click()
    frmActions.Show
End Sub
Private Sub btnAddClient_Click()
    frmClients.Show
End Sub
Private Sub btnClear_Click()
    Call clearAll(myTextboxes)
    txtDateDeReceptionShow.Text = ""
    With txtEssaiSortiLeDateShow
        .Locked = False
        .Text = ""
        .Locked = True
    End With
    lblDemandeurDetails.Caption = ""
    lblPayeurDetails.Caption = ""
    lblEDemandeurDetails.Caption = ""
    lblEPayeurDetails.Caption = ""
    lblFicheResultatsStatus.Caption = ""
    chkLegacy.Value = False
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnEssaiTypeHelp_Click()
MsgBox _
"EM = Essais Métalliques (boulons, torons, soudure...)" & vbCrLf & _
"ENM = Essais Non Métalliques (béton, briques, ciments...)" & vbCrLf & _
"ES = Essais Spéciaux mécaniques, sans normes (tarages...)" & vbCrLf & _
"R = Recherches" _
, Buttons:=vbQuestion, Title:="Types d'Essais"
End Sub

Private Sub btnRepeter_Click()
'Get a new ID
    txtEssaiID.Text = latestEssaiID() + 1
'Clear version and sorti dates
    txtEssaiSortiLeDateShow.Text = ""
    txtEssaiSortiLeDate.Text = ""
    txtEssaiVersion = ""
    lblFicheResultatsStatus.Caption = ""
'Make textboxes enabled
    Dim itxt As Variant
    For Each itxt In myTextboxes
        itxt.Locked = False
    Next itxt
End Sub

Private Sub btnSearch_Click()

    If chkLegacy.Value = False Then
'Normal search
        Call searchTable("Essais", "essaisTable", myTextboxes)
        
        'Show clients details
        Call fillInAllClientDetails
        
    Else
'Do a search in LEGACY workbook
        Dim legacyEssaisBook As Workbook: Dim legacyClientsBook As Workbook
        
        Set legacyEssaisBook = Workbooks.Open(LEGACYWORKBOOKSFOLDER & "\" & LEGACYESSAISFILENAME)
        Call searchTable("Essais", "essaisTable", myTextboxes)
        
        'make textboxes readonly - non-modifiable
        Dim itxt As Variant
        For Each itxt In myTextboxes
            itxt.Locked = True
        Next itxt
        MsgBox "Les essais legacy sont non-modifiable et les champs ont été désactivés. Appuyer sur Vider pour réinitialiser le formulaire.", Buttons:=vbInformation, Title:="Essai Legacy"
        
        'close legacy workbook
        legacyEssaisBook.Close
        
        'Show clients details from legacy workbook
        Set legacyClientsBook = Workbooks.Open(LEGACYWORKBOOKSFOLDER & "\" & LEGACYCLIENTSFILENAME)
        Call fillInAllClientDetails
        legacyClientsBook.Close
        
    End If
    
    'Dates are stored and accessed as numbers, but are shown to user as a date
    'This is so the format doesn't mess up e.g. 8/7/18 becoming 7/8/18
    txtDateDeReceptionShow.Text = txtDateDeReception.Text
    txtEssaiSortiLeDateShow.Text = txtEssaiSortiLeDate.Text
    If IsDate(txtDateDeReception.Text) Then txtDateDeReception.Text = Int(CDbl(CDate(txtDateDeReception.Text)))
    If IsDate(txtEssaiSortiLeDate.Text) Then txtEssaiSortiLeDate.Text = Int(CDbl(CDate(txtEssaiSortiLeDate.Text)))
    
    'If there is an associated Fiche Resultats, notify that one exists
    If Dir(CreateDocumentFileName("RE", txtEssaiID.Text, docext:=".xls")) <> "" Then
        lblFicheResultatsStatus.Caption = "Fiche resultats lié a cet essai"
    Else
        lblFicheResultatsStatus.Caption = ""
    End If
    
    'If there is a comment on the version, show it
    Dim myIDCell As Range: Set myIDCell = findInColumn(txtEssaiID.Text, COLUMNESSAIID, "Essais", "essaisTable")
    If Not myIDCell Is Nothing Then
        If Not myIDCell.EntireRow.Cells(1, COLUMNESSAIVERSION).Comment Is Nothing Then
            MsgBox "Commentaire ajouté de la version " & myIDCell.EntireRow.Cells(1, COLUMNESSAIVERSION).Value & ":" & vbCrLf & myIDCell.EntireRow.Cells(1, COLUMNESSAIVERSION).Comment.Text _
            , Buttons:=vbInformation, Title:="Commentaire de la version"
        End If
    End If
    
End Sub

Public Sub fillInClientDetails(client As String)
'Fill associated details label with client details associated with client's ID
'Textbox and Label names should be in the correct format e.g. txtDemandeurID and lblDemandeurDetails
    Dim IDsource As Variant: Dim detailsLabel As Variant
    Dim IDcell As Range
    
    Set detailsLabel = Me.Controls("lbl" & client & "Details")
    Set IDsource = Me.Controls("txt" & client & "ID")
    
    If IDsource.Text <> "" Then
        Set IDcell = findInColumn(IDsource.Text, COLUMNCLIENTID, "Clients", "clientsTable")
        
        'Fetch client details from ID cell
        If Not IDcell Is Nothing Then _
            detailsLabel.Caption = concatenateRow(IDcell.EntireRow, 2)
    End If
End Sub

Public Sub fillInAllClientDetails()
    Call fillInClientDetails("Demandeur")
    Call fillInClientDetails("Payeur")
    Call fillInClientDetails("EDemandeur")
    Call fillInClientDetails("EPayeur")
End Sub
Private Sub btnDatePicker_Click()
'Get today's date and display as date to user, but as integer for saving
    txtDateDeReception.Text = Int(CDbl(Now()))
    txtDateDeReceptionShow.Text = CStr(Format(Now(), DATEFORMAT))
End Sub

Private Sub cbxEssaiType_AfterUpdate()
'if changed to ENM then what material? Default/cubes/carottes/cylindres?
    If cbxEssaiType.Text = "ENM" Then
        Dim materialInput As String
        materialInput = InputBox("Choisissez type de formulaire ENM " & vbCrLf & _
            "0 - Default" & vbCrLf & _
            "1 - Cubes" & vbCrLf & _
            "2 - Carottes" & vbCrLf & _
            "3 - Cylindres" _
            , "ENM", Default:="0")
        
        Dim natureText As String: natureText = ""
        
        Select Case materialInput
            Case "1"
                natureText = " cube(s) en béton, dimensions déclarées: "
            Case "2"
                natureText = " carotte(s) en béton "
            Case "3"
                natureText = " cylindre(s) en béton, dimensions nominales: "
            Case Else
                Exit Sub
        End Select
        
    'Add this additional info to nature du produit
    txtNatureDuProduit.Text = natureText
        
    End If
End Sub


Private Sub txtDateDeReceptionShow_AfterUpdate()
'Two fields are needed, one as integer and one as date for user, as VBA built-in date locale is messed up
'txtDateDeReceptionShow is used purely for entry, it is not saved anywhere
'Note e.g. "8 mar 18" is valid too as input

    If IsDate(txtDateDeReceptionShow.Text) Then
        'Format as date for user
        txtDateDeReceptionShow.Text = Format(CDate(txtDateDeReceptionShow.Text), DATEFORMAT)
        'Format as integer for saving to table
        txtDateDeReception.Text = Int(CDbl(CDate(txtDateDeReceptionShow.Text)))
    
    ElseIf txtDateDeReceptionShow.Text = "" Then
    Else
        MsgBox "Invalid date", vbExclamation, "Invalid Date"
        txtDateDeReceptionShow.Text = ""
    End If
End Sub

Private Sub UserForm_Activate()

    initialiseFilePaths
    
    'Populate comboboxes
    With cbxEssaiType
        .AddItem "EM"
        .AddItem "ENM"
        .AddItem "ES"
        .AddItem "R"
    End With
    
    With cbxNorme
        .AddItem "NBN EN 12390-3"
        .AddItem "NBN EN ISO 15630-1"
        .AddItem "NBN EN ISO 15630-2"
        .AddItem "NBN EN ISO 15630-3"
    End With
    
    With cbxEssaiAccredite
        .AddItem "Non"
        .AddItem "Oui"
    End With
    
    
    'Collection of all textboxes on page.
    'NOTE: this must be in the same order as column headers in the table
    'NOTE: update the constants in tableManagement with the correct column numbers
    With myTextboxes
    .Add txtEssaiID
    .Add cbxEssaiType
    .Add txtEssaiVersion
    .Add txtEssaiSortiLeDate
    .Add cbxEssaiAccredite
    .Add txtDemandeurID
    .Add txtPayeurID
    .Add txtEDemandeurID
    .Add txtEPayeurID
    .Add txtReferences
    .Add txtQuantity
    .Add txtNatureDuProduit
    .Add txtDateDeReception
    .Add txtProvenance
    .Add txtEssaisDemandes
    .Add cbxNorme
    .Add txtRemarques
    .Add txtTechnicien
    .Add txtAutreCoordonnees
    End With
    
    Load frmActions
    frmActions.MultiPage1.Pages("pgeSortir").Visible = False
    frmActions.Show
    
End Sub

Public Sub saveEssai(Optional modrow As Range = Nothing, Optional rowpos As Range = Nothing)
    If modrow Is Nothing And rowpos Is Nothing Then
        Call saveToTable("Essais", "essaisTable", myTextboxes)
    ElseIf rowpos Is Nothing Then
        Call saveToTable("Essais", "essaisTable", myTextboxes, modifyrow:=modrow)
    Else
        Call saveToTable("Essais", "essaisTable", myTextboxes, newRowAfter:=rowpos)
    End If
End Sub

Public Function modsMade() As Boolean
'Returns True if boxes on form are any different from corresponding row in table
    
    Dim table As ListObject
    Set table = Worksheets("Essais").ListObjects("essaisTable")
    Dim myRow As Range
    Set myRow = findInColumn(txtEssaiID.Text, COLUMNESSAIID, "Essais", "essaisTable")
    
    modsMade = False
    
    Dim i As Integer
    For i = 1 To myTextboxes.count
        If CStr(myRow.Cells(1, i).Value2) <> CStr(myTextboxes(i).Text) Then
            'Modification made somewhere!
            modsMade = True
            Exit For
        End If
    Next i
End Function
