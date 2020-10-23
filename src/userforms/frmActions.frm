VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmActions 
   Caption         =   "Actions"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmActions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   LOGICIELGC 2018 frmActions   '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnConsult_Click()
    'display instructions
    MsgBox "Veuillez remplir au moins un blanc, puis appuyer sur Rechercher pour charger des entrées", Buttons:=vbInformation, Title:="Consulter"
    Unload Me
End Sub

Private Sub btnDelete_Click()
'NOTE: doesn't actually delete ID, or else things might get messy
'Replaces all essai fields with supprimé
    Dim essaiToDelete As String
    Dim sure As Variant
    
    essaiToDelete = InputBox("Saisir essai à supprimer", "Supprimer", Default:=frmEssais.txtEssaiID.Text)
    If essaiToDelete = "" Then Exit Sub
            
    Dim targetID As Range
    Dim rowToDelete As Range
    
    'find row to delete
    Set targetID = findInColumn(essaiToDelete, COLUMNESSAIID, "Essais", "essaisTable")
    
    If targetID Is Nothing Then
        MsgBox "Essai n'existe pas!", Buttons:=vbExclamation, Title:="Supprimer"
        Exit Sub
    End If
    
    sure = MsgBox("Vous êtes sur?", vbOKCancel + vbExclamation, "Supprimer")
    If sure = vbOK Then
        'delete row!
        Set rowToDelete = targetID.EntireRow
        
        If chkPreserverID.Value = True Then
        'Keep row, but get rid of all details
            Dim i As Integer: i = 2
            Do While cellInTable(rowToDelete.Cells(1, i)) = True
                rowToDelete.Cells(1, i).Value = "supprimé"
                i = i + 1
            Loop
        Else
        'Delete entire row
            rowToDelete.Delete
        End If
        
        MsgBox "Success!", Buttons:=vbInformation, Title:="Supprimer"
    End If
End Sub

Private Sub btnEntrerResultats_Click()
    Dim essaiToOpen As String
    
    essaiToOpen = InputBox("Saisir essai à ouvrir", "Entrer", Default:=frmEssais.txtEssaiID.Text)
    If essaiToOpen = "" Then Exit Sub
            
    Dim targetID As Range
    Set targetID = findInColumn(essaiToOpen, COLUMNESSAIID, "Essais", "essaisTable")
    Call OpenResultsFile(targetID)
End Sub

Sub OpenResultsFile(cellIDToOpen As Range)
'Open or create new results for fiche essai, and format table based on the id type/material
    
    'Setup opening workbook
    If cellIDToOpen Is Nothing Then
        MsgBox "ID n'existe pas.", vbExclamation, "Fiche Resultats"
        Exit Sub
    End If
    
    Dim saveFileName As String
    Dim myID As String
    myID = cellIDToOpen.EntireRow.Cells(1, COLUMNESSAIID).Text
    saveFileName = CreateDocumentFileName("RE", myID, docext:=".xls")
    
    'Try to open results workbook for ID
    On Error GoTo CantOpen
    Workbooks.Open saveFileName
    
    Exit Sub
    
CantOpen:
    'Results file doesn't exist. Create new?
    If MsgBox("Fiche de resultats n'existe pas. Créer nouveau?", Buttons:=vbInformation + vbYesNo, Title:="Fiche Resultats") _
        = vbNo Then Exit Sub
    
    
'To add a new template, see fileManagement module
    
    'Select a template
    Load frmMultipleResults
    
    frmMultipleResults.lblDescription.Caption = "Choisissez un masque pour le type d'essai..."
    With frmMultipleResults.lbxMultipleResults
        .ColumnCount = 3
        .ColumnWidths = "50;50;50"
        
        'Add all templates to multicolumn listbox
        Dim j As Integer
        For j = 0 To UBound(REtemplateArray, 1) - 1
            .AddItem REtemplateArray(j)(1): .List(j, 1) = REtemplateArray(j)(2): .List(j, 2) = REtemplateArray(j)(3)
        Next j
        
    End With
    
    frmMultipleResults.Show
    
    'Default variables that can be changed based on chosen format
    Dim templateName As String: templateName = "TemplateResultatsDefault.xlt"   'Default template to open
    Dim ResultatRow1 As Integer: ResultatRow1 = 16                               'Default first row of table

    'Get template details based on selection
    templateName = REtemplateArray(selectedIndex)(0)
    ResultatRow1 = REtemplateArray(selectedIndex)(4)
    
    Dim q As String: q = cellIDToOpen.EntireRow.Cells(1, COLUMNESSAIQUANTITY).Text
    Dim quantity As Integer: quantity = 0
    If IsNumeric(q) = True Then quantity = CInt(q)
    
    'Open template
    Dim newWb As Workbook
    Set newWb = Workbooks.Add(Template:=TEMPLATESFOLDER & "\" & templateName)
    
    'Add details
    eRT newWb.Worksheets(1), "<<EssaiID>>", cellIDToOpen.EntireRow.Cells(1, COLUMNESSAITYPE).Text & " " & cellIDToOpen.EntireRow.Cells(1, COLUMNESSAIID).Text
    eRT newWb.Worksheets(1), "<<essaisDemandes>>", cellIDToOpen.EntireRow.Cells(1, COLUMNESSAISDEMANDES).Text
    Dim normeText As String: normeText = cellIDToOpen.EntireRow.Cells(1, COLUMNESSAINORME).Text
    If normeText = "" Then normeText = "N/A"
    eRT newWb.Worksheets(1), "<<Norme>>", normeText
    'eRT newWb.Worksheets(1), "<<Remarques>>", cellIDToOpen.EntireRow.Cells(1, COLUMNESSAIREMARQUES).Text

    'Add new rows to table
    If quantity > 1 Then
        Dim i As Integer
        For i = 1 To quantity - 1
            newWb.Worksheets(1).Rows(ResultatRow1).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
        Next i
    End If
    
    'Save with ID
    newWb.CheckCompatibility = False
    newWb.SaveAs fileName:=saveFileName, FileFormat:=xlExcel8
    
    selectedIndex = 0
    Exit Sub
    
End Sub

Private Sub btnSortir_Click()
'Create new FV/PG using word template, then replace <<>> placeholders with textbox data
    Dim sortirType As String
    Dim templateName As String
    Dim saveFileName As String
    
    'Read which document we want
    Select Case cbxSortir.Text
    Case "Fiche Verte"
        templateName = FICHEVERTETEMPLATE
        sortirType = "FV"
    Case "Page de Garde"
        templateName = PAGEDEGARDETEMPLATE
        sortirType = "PG"
    Case Else
        MsgBox "Selectionner fiche à sortir", vbExclamation, "Sortir"
        Exit Sub
    End Select
    
    'Save current frmEssais and create new version if needed
    'Make sure that details are there!
    If frmEssais.txtDemandeurID.Text = "" Then
        MsgBox "Rechercher ou remplir plus de blancs avant d'enregistrer l'essai.", vbExclamation, "Enregistrer"
        Exit Sub
    End If
    
    Call btnSave_Click
    
    'Set up to start creating documents
    Dim myID As String: myID = frmEssais.txtEssaiID.Text
    If myID = "" Then
        MsgBox "Essai ID not entered", vbExclamation, "Sortir"
        Exit Sub
    End If
    Dim myIDCell As Range: Set myIDCell = findInColumn(myID, COLUMNESSAIID, "Essais", "essaisTable")
    
    If chkImprimer.Value = True And sortirType = "FV" Then _
        MsgBox "Mettre des feuilles vertes dans l'imprimante", Buttons:=vbExclamation, Title:="Imprimer"
    
    'Create Word document objects
    Dim wApp As Object
    Dim wDoc As Object
    
    Set wApp = CreateObject("Word.Application")
    
    'if fv then carry on (has no versions)
    'for page de garde, add sorti date if there isn't one already - this means it's now sorti
    If sortirType = "PG" Then
        'If myIDCell.EntireRow.Cells(1, COLUMNESSAISORTILEDATE).Value = "" Then
            myIDCell.EntireRow.Cells(1, COLUMNESSAISORTILEDATE).Value = Int(CDbl(Now()))
    End If
    
    'Retrieve version number
    Dim versionNo As Integer
    versionNo = myIDCell.EntireRow.Cells(1, COLUMNESSAIVERSION).Value
    
    'Retrieve accreditation status
    Dim accreditedStatus As Boolean: accreditedStatus = False
    'If frmEssais.cbxNorme.ListIndex <> -1 Then accreditedStatus = True
    If LCase(frmEssais.cbxEssaiAccredite.Text) = "oui" Then accreditedStatus = True
    
    'create document based on template
    Set wDoc = wApp.Documents.Add(Template:=TEMPLATESFOLDER & "\" & templateName, NewTemplate:=False, DocumentType:=0)
    
    'Fill in template
    On Error Resume Next
    With frmEssais
    Dim clientRow As Range
    'firstly, fields in both templates
        wRT wDoc, "<<EssaiType>>", .cbxEssaiType.Text
        wRT wDoc, "<<EssaiID>>", myID
        
        Set clientRow = findInColumn(.txtDemandeurID.Text, COLUMNCLIENTID, "clients", "clientstable").EntireRow
        wRT wDoc, "<<DemandeurID>>", .txtDemandeurID.Text
        wRT wDoc, "<<DemandeurNom>>", clientRow.Cells(1, COLUMNCLIENTNOM)
        wRT wDoc, "<<DemandeurAdresse>>", clientRow.Cells(1, COLUMNCLIENTADRESSE)
        
        Set clientRow = findInColumn(.txtPayeurID.Text, COLUMNCLIENTID, "clients", "clientstable").EntireRow
        wRT wDoc, "<<PayeurID>>", .txtPayeurID.Text
        wRT wDoc, "<<PayeurNom>>", clientRow.Cells(1, COLUMNCLIENTNOM)
        wRT wDoc, "<<PayeurAdresse>>", clientRow.Cells(1, COLUMNCLIENTADRESSE)
        
        wRT wDoc, "<<References>>", .txtReferences.Text
        wRT wDoc, "<<NatureDuProduit>>", .txtQuantity.Text & " " & .txtNatureDuProduit.Text
        wRT wDoc, "<<Provenance>>", .txtProvenance.Text
        wRT wDoc, "<<DateDeReception>>", Format(CDate(.txtDateDeReception.Text), DATEFORMAT)
        wRT wDoc, "<<EssaisDemandes>>", .txtEssaisDemandes.Text
        wRT wDoc, "<<Remarques>>", .txtRemarques.Text

        'Deal with accreditation
        Dim normeText As String: normeText = "": Dim accredText As String: accredText = ""
        Dim BELACLogo As String: BELACLogo = ""
        If .cbxNorme.Text <> "" Then normeText = "selon " & .cbxNorme.Text
        
        If accreditedStatus = True Then
            'Yes, an accredited norme is chosen
            accredText = "Essai accrédité certificat B392-Test : Voir annexe"
            Worksheets("Utilities").Shapes("logoBELAC").Copy
            BELACLogo = "^c"
        Else
            'No, not accredited.
        End If
        
        wRT wDoc, "<<logoBELAC>>", BELACLogo
        wRT wDoc, "<<Norme>>", normeText
        wRT wDoc, "<<PhraseAccreditation>>", accredText
        
        
    'secondly, fields only in fiche verte
        If sortirType = "FV" Then
            If .txtEDemandeurID.Text <> "" Then
            Set clientRow = findInColumn(.txtEDemandeurID.Text, COLUMNCLIENTID, "clients", "clientstable").EntireRow
            wRT wDoc, "<<EDemandeurID>>", .txtEDemandeurID.Text
            wRT wDoc, "<<EDemandeurNom>>", clientRow.Cells(1, COLUMNCLIENTNOM)
            wRT wDoc, "<<EDemandeurAdresse>>", clientRow.Cells(1, COLUMNCLIENTADRESSE)
            End If
            
            If .txtEPayeurID.Text <> "" Then
            Set clientRow = findInColumn(.txtEPayeurID.Text, COLUMNCLIENTID, "clients", "clientstable").EntireRow
            wRT wDoc, "<<EPayeurID>>", .txtEPayeurID.Text
            wRT wDoc, "<<EPayeurNom>>", clientRow.Cells(1, COLUMNCLIENTNOM)
            wRT wDoc, "<<EPayeurAdresse>>", clientRow.Cells(1, COLUMNCLIENTADRESSE)
            End If
            
            wRT wDoc, "<<DateModifie>>", Format(Now(), DATEFORMAT)
            wRT wDoc, "<<AutresCoordonnees>>", .txtAutreCoordonnees.Text
    
    'thirdly, fields only in page de garde
        Else
            'Essai sorti le date
            wRT wDoc, "<<EssaiSortiLeDate>>", Format(CDate(myIDCell.EntireRow.Cells(1, COLUMNESSAISORTILEDATE).Value2), DATEFORMAT)
            
            'Version - if more than 1, show version, and show old version date too
            Dim versionText As String: versionText = ""
            Dim previousVersionRow As Range
            Dim annuleText As String: annuleText = ""
            
            If versionNo > 1 Then
                versionText = "v" & CStr(versionNo)
                Set previousVersionRow = myIDCell.EntireRow.Offset(-1, 0)
                annuleText = "Annule et remplace la version " & previousVersionRow.Cells(1, COLUMNESSAIVERSION) _
                    & " sorti le " & previousVersionRow.Cells(1, COLUMNESSAISORTILEDATE)
            End If
            
            wRT wDoc, "<<EssaiVersion>>", versionText
            wRT wDoc, "<<AnnuleText>>", annuleText
            
            'Accreditation annexes
            If accreditedStatus = False _
                Then Call wordRemovePagesFromEnd(wApp, wDoc, 2) 'remove certificate+scope
            
            'Excel results
            Dim resultatsFileName As String: resultatsFileName = CreateDocumentFileName("RE", myID, docext:=".xls")
            Dim voirExcelText As String
            'voirExcelText = "La fiche de résultats liée à cet essai " & resultatsFileName & " est reprise dans la/les page(s) suivante(s)."
            
            wRT wDoc, "<<voirExcelText>>", " ", moveSelection:=True
            
            'Add results table
            On Error Resume Next
            Dim excelShape As Variant
            If Dir(resultatsFileName) <> "" Then _
            Set excelShape = wDoc.InlineShapes.AddOLEObject(ClassType:="Excel.Sheet.8", fileName:= _
                resultatsFileName, LinkToFile:=True, DisplayAsIcon:=False, Range:=wApp.Selection.Range)
            
            'Reshape table
            Dim percent As Variant
            percent = wDoc.PageSetup.TextColumns.Width / excelShape.Width
            excelShape.ScaleWidth = 100 * percent
            excelShape.ScaleHeight = 100 * percent
            
            
        End If
    End With
    
    saveFileName = CreateDocumentFileName(sortirType, myID, docversion:=versionNo)
    wDoc.SaveAs fileName:=(saveFileName), FileFormat:=0, AddtoRecentFiles:=False
    
    'Open fiche essai excel file if wanted (for fiche verte)
    If chkOuvrirAvecFicheEssai.Value = True And cbxSortir.Text = "Fiche Verte" Then
        Call OpenResultsFile(myIDCell)
    End If
    
    'Print if needed
    If chkImprimer.Value = True Then
        wDoc.PrintOut
    End If
    
    wApp.Visible = True 'Show word doc
    Unload Me: Unload frmEssais
    Exit Sub
    
End Sub


Private Sub btnNew_Click()
    'display instructions
    MsgBox "Pour créer une entrée à partie d'une entrée existante, veuillez utiliser Rechercher", Buttons:=vbInformation, Title:="Nouveau"
    'create new ID and input
    frmEssais.txtEssaiID.Text = latestEssaiID() + 1
    'input default fields
    frmEssais.txtTechnicien.Text = "Carl Vroomen"
    Unload Me
End Sub


Private Sub btnSave_Click()
   
    With frmEssais
        If .txtDemandeurID.Text = "" Then
            MsgBox "Rechercher ou remplir plus de blancs avant d'enregistrer l'essai.", vbExclamation, "Enregistrer"
            Exit Sub
        End If
            
        Dim myID As Range
        Set myID = findInColumn(.txtEssaiID.Text, COLUMNESSAIID, "Essais", "essaisTable")
        
        'myID should be latest version
        
        'Is the ID already there?
        If myID Is Nothing Then
            'No: save a new row
            Call .saveEssai
        Else
        
            'Yes: is it actually worth doing anything?
            If .modsMade() = False Then Exit Sub
            
            'Retrieve essai version
            Dim myIDVersion As String
            myIDVersion = myID.EntireRow.Cells(1, COLUMNESSAIVERSION).Value
            
            'Is it already sorti? / even if it is do we want a new version?
            If myID.EntireRow.Cells(1, COLUMNESSAISORTILEDATE).Value = "" Then
                'No: just modify entry
                MsgBox "Modification...", Buttons:=vbInformation
                Call .saveEssai(modrow:=myID.EntireRow)
                
            ElseIf MsgBox("La page de garde a été déjà sorti. Créer une nouvelle version de l'essai?" & vbCrLf & "Sinon, la modification s'effectuera sans créer un /2 ou /3 etc." & vbCrLf & _
                "Répondre oui si le rapport a été déjà envoyé au client.", Buttons:=vbExclamation + vbYesNo, Title:="Sortir") = vbNo Then
                
                'No: just modify entry
                MsgBox "Modification...", Buttons:=vbInformation
                Call .saveEssai(modrow:=myID.EntireRow)
            
            Else
                'Yes: add ! to myID, then create new row with new version
                
                'Add ! to current myID
                myID.Value = "!" & myID.Value
                
                'Make current one version 1 if nothing there
                If myIDVersion = "" Then
                    myIDVersion = 1
                    myID.EntireRow.Cells(1, COLUMNESSAIVERSION).Value = 1
                End If
                              
                'Create a new version
                .txtEssaiVersion = CInt(myIDVersion) + 1
                .txtEssaiSortiLeDate = ""
                .txtEssaiSortiLeDateShow = ""
                Call .saveEssai(rowpos:=myID.EntireRow)
                
                'Call .fillInAllClientDetails 'just in case
                
                'Add comment to current one saying why new created
                Dim versionComment As String
                versionComment = InputBox("Ajouter un commentaire à propos de la version?", "Enregistrer")
                If versionComment <> "" Then myID.EntireRow.Offset(1, 0).Cells(1, COLUMNESSAIVERSION).AddComment versionComment
            
            End If

        End If
    End With
    
    Call updateDernierNumero
End Sub

'NOT NEEDED: if you want to create it again just sortir should work, or else find it yourself
Private Sub btnView_Click()
    Dim viewFileName As String
    Dim viewType As String

    'Read which document we want
    Select Case cbxSortir.Text
    Case "Fiche Verte"
        viewType = "FV"
    Case "Page de Garde"
        viewType = "PG"
    Case Else
        MsgBox "Selectionner fiche à inspecter", vbExclamation, "Inspecter"
        Exit Sub
    End Select

    'Set up to start opening documents
    Dim myID As String: myID = frmEssais.txtEssaiID.Text
    If myID = "" Then
        MsgBox "Essai ID not entered", vbExclamation, "Inspecter"
        Exit Sub
    End If
    Dim myIDCell As Range: Set myIDCell = findInColumn(myID, COLUMNESSAIID, "Essais", "essaisTable")
    If myIDCell Is Nothing Then Exit Sub
    
    Dim wApp As Object
    Set wApp = CreateObject("Word.Application")
    
    Dim versionNo As Integer
    versionNo = myIDCell.EntireRow.Cells(1, COLUMNESSAIVERSION).Value
    
    'Make filename to open
    viewFileName = CreateDocumentFileName(viewType, myID, docversion:=versionNo)
    
    On Error GoTo CantOpen
    wApp.Documents.Open viewFileName

    'Open fiche essai excel file if wanted
    If chkOuvrirAvecFicheEssai.Value = True And cbxSortir.Text = "Fiche Verte" Then
        Call OpenResultsFile(myIDCell)
    End If
    
    wApp.Visible = True 'open word doc
    Unload Me: Unload frmEssais
    
    Exit Sub
CantOpen:
    MsgBox "N'a pas pu ouvrir le document. Vérifier que le document a été sorti.", vbCritical, "Inspecter"
End Sub

Private Sub cbxSortir_Change()
    If cbxSortir.Text = "Fiche Verte" Then
        btnSortir.BackColor = RGB(0, 255, 0)
        btnView.BackColor = RGB(0, 255, 0)
        chkOuvrirAvecFicheEssai.Visible = True
    Else
        btnSortir.BackColor = btnSave.BackColor
        btnView.BackColor = btnSave.BackColor
        chkOuvrirAvecFicheEssai.Visible = False
    End If
End Sub

Private Sub btnShowVersions_Click()
    With frmEssais
        If .txtEssaiID.Text = "" Then Exit Sub
        
        Worksheets("Essais").ListObjects("essaisTable").Range.AutoFilter Field:=COLUMNESSAIID, _
            Criteria1:="=" & .txtEssaiID.Text, Operator:=xlOr, _
            Criteria2:="=*" & .txtEssaiID.Text
    End With
    Unload Me
    Unload frmEssais
End Sub


Private Sub UserForm_Initialize()
    With cbxSortir
        .AddItem "Fiche Verte"
        .AddItem "Page de Garde"
    End With
    Call updateDernierNumero
End Sub
Sub updateDernierNumero()
lblDernierPVNumero.Caption = "Dernier PV numéro " & latestEssaiID
End Sub
