VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalcul 
   Caption         =   "Calcul de la date d'essai des bétons"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3705
   OleObjectBlob   =   "frmCalcul.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalcul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   LOGICIELGC 2018 frmCalculs   '
'   Andrew Wang, September 2018  '
'                                '
'     andrew.wang27gmail.com     '
'                                '

Private Sub btnClear_Click()
    txtDateEssai.Text = ""
    txtDateFabrication.Text = ""
    txtAge.Text = ""
    lblDayEssai.Caption = ""
    lblDayFabrication.Caption = ""
End Sub

Private Sub btnDatePicker_Click()
'Get today's date
txtDateFabrication.Text = Format(Now(), "DD/MM/YYYY")
txtDateFabrication.SetFocus
End Sub


Private Sub txtDateEssai_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If txtDateEssai.Text <> "" Then
    txtDateEssai.Text = Format(CDate(txtDateEssai.Text), "dd/mm/yyyy")
    lblDayEssai.Caption = Format(CDate(txtDateEssai.Text), "dddd")
    Call Calculate
Else
    lblDayEssai.Caption = ""
End If
End Sub

Private Sub txtDateFabrication_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If txtDateFabrication.Text <> "" Then
    txtDateFabrication.Text = Format(CDate(txtDateFabrication.Text), "dd/mm/yyyy")
    lblDayFabrication.Caption = Format(CDate(txtDateFabrication.Text), "dddd")
    Call Calculate
Else
    lblDayFabrication.Caption = ""
End If
End Sub

Private Sub txtAge_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Call Calculate
End Sub


Public Sub Calculate()
    Dim tdf As String: tdf = txtDateFabrication.Text
    Dim ta As String: ta = txtAge.Text
    Dim tde As String: tde = txtDateEssai.Text
    
    If tdf <> "" And ta <> "" Then tde = CDate(tdf) + CInt(ta)
    If tdf <> "" And tde <> "" Then ta = CDate(tde) - CDate(tdf)
    If tde <> "" And ta <> "" Then tdf = CDate(tde) - CInt(ta)
    
    txtDateFabrication.Text = tdf
    txtAge.Text = ta
    txtDateEssai.Text = tde
    
End Sub
