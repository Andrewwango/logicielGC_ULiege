VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClients 
   Caption         =   "Add Clients..."
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5010
   OleObjectBlob   =   "frmClients.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   LOGICIELGC 2018 frmClients   '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '

Dim myTextboxes As New Collection

'Display names for combobox
Const DemandeurDisplayName As String = "Demandeur"
Const PayeurDisplayName As String = "Payeur"
Const EDemandeurDisplayName As String = "Demandeur Expeditation"
Const EPayeurDisplayName As String = "Payeur Expeditation"


Private Sub btnAddClient_Click()
'Populate essai form with client ID and details
'Note this button only appears when this form is opened from frmEssais
    
    Dim legacyClientsBook As Workbook
    If chkLegacy.Value = True Then Set legacyClientsBook = Workbooks.Open(LEGACYWORKBOOKSFOLDER & "\" & LEGACYCLIENTSFILENAME)

    Select Case cbxClientType.Text
        Case DemandeurDisplayName
            frmEssais.txtDemandeurID = frmClients.txtClientID
            frmEssais.txtRemarques = frmClients.txtClientRemarques 'Demandeur remarques are important!
            frmEssais.txtAutreCoordonnees = frmClients.txtClientAutre 'Add any additional coordonnees!
            Call frmEssais.fillInClientDetails("Demandeur")
            
        Case PayeurDisplayName
            frmEssais.txtPayeurID = frmClients.txtClientID
            Call frmEssais.fillInClientDetails("Payeur")
        Case EDemandeurDisplayName
            frmEssais.txtEDemandeurID = frmClients.txtClientID
            Call frmEssais.fillInClientDetails("EDemandeur")
        Case EPayeurDisplayName
            frmEssais.txtEPayeurID = frmClients.txtClientID
            Call frmEssais.fillInClientDetails("EPayeur")
        
        Case Else
            MsgBox "Choisissez un client à ajouter", Buttons:=vbExclamation, Title:="Add Clients"
            Exit Sub
    End Select
    
    If chkLegacy.Value = True Then legacyClientsBook.Close
    
    Call clearAll(myTextboxes)
End Sub

Private Sub btnClear_Click()
    Call clearAll(myTextboxes)
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnIdem_Click()
    'Copies clientID from existing and input automatically
    Select Case cbxClientType.Value
        Case PayeurDisplayName, EDemandeurDisplayName
            frmClients.txtClientID = frmEssais.txtDemandeurID.Text
        Case EPayeurDisplayName
            frmClients.txtClientID = frmEssais.txtPayeurID.Text
    End Select
    
    Call btnAddClient_Click
End Sub

Private Sub btnSave_Click()
'Save new client, or modify existing
    
    'Check whether new automatic ID wanted or not
    If txtClientID.Text = "" Then
        'Create new ID according to name (e.g. W1000 becomes W1001)
        Dim firstChar As String
        Dim latestID As String
        
        firstChar = UCase(Left(txtClientNom.Text, 1))
        latestID = latestClientID(firstChar)
        
        txtClientID.Text = firstChar & Format$(CStr(CInt(Right(latestID, Len(latestID) - 1)) + 1), "0000")
    Else
        'Check if ID already exists
        Dim myRow As Range
        Set myRow = findInColumn(txtClientID.Text, COLUMNCLIENTID, "Clients", "clientsTable")
        
        If Not myRow Is Nothing Then
            MsgBox "ClientID existe déja, modification...", Buttons:=vbInformation
            Call saveToTable("Clients", "clientsTable", myTextboxes, modifyrow:=myRow.EntireRow)
            Exit Sub
        End If
    End If
    
    'All done, save a new client according to textboxes on form
    Call saveToTable("Clients", "clientsTable", myTextboxes)
End Sub

Private Sub btnSearch_Click()
    
    If chkLegacy.Value = False Then
        Call searchTable("Clients", "clientsTable", myTextboxes)
    Else
        Dim legacyClientsBook As Workbook
        Set legacyClientsBook = Workbooks.Open(LEGACYWORKBOOKSFOLDER & "\" & LEGACYCLIENTSFILENAME)
        
        Call searchTable("Clients", "clientsTable", myTextboxes)
        legacyClientsBook.Close
    End If
End Sub

Private Sub cbxClientType_Change()
'Display Idem button when needed!
    Select Case cbxClientType.Value
        Case DemandeurDisplayName
            btnIdem.Visible = False
        Case PayeurDisplayName, EDemandeurDisplayName, EPayeurDisplayName
            btnIdem.Visible = True
    End Select
End Sub

Private Sub UserForm_Initialize()

    initialiseFilePaths
    
    With cbxClientType
        .AddItem (DemandeurDisplayName)
        .AddItem (EDemandeurDisplayName)
        .AddItem (PayeurDisplayName)
        .AddItem (EPayeurDisplayName)
    End With
    
    'Collection of all textboxes on page.
    'NOTE: this must be in the same order as column headers in the table
    'NOTE: update the constants in tableManagement with the correct column numbers
    
    With myTextboxes
    .Add txtClientID
    .Add txtClientNom
    .Add txtClientAdresse
    .Add txtClientAutre
    .Add txtClientRemarques
    End With
End Sub
