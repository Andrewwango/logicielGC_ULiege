VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMultipleResults 
   Caption         =   "Choisissez..."
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8565
   OleObjectBlob   =   "frmMultipleResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMultipleResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   LOGICIELGC 2018 frmMultipleResults   '
'   Andrew Wang, September 2018          '
'                                        '
'     andrew.wang27gmail.com             '
'                                        '

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnDone_Click()
    
    If lbxMultipleResults.Value <> "" Then
        selectedRow = lbxMultipleResults.Value
        selectedIndex = lbxMultipleResults.ListIndex
    Else
        selectedRow = ""
        selectedIndex = 0
    End If
    
    Unload Me
End Sub
