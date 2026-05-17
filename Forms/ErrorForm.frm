VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorForm
   Caption         =   "UserForm1"
   ClientHeight    =   2490
   ClientLeft      =   420
   ClientTop       =   1620
   ClientWidth     =   11400
   OleObjectBlob   =   "ErrorForm.frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Sub OK_Click()
    Unload ErrorForm
End Sub

Private Sub UserForm_Initialize()
    Dim hwnd As LongPtr

    With ErrorForm
	.Height = 160.25        'fix screen driver issue
	.Width = 568.25

	.WhiteBg.Width = .Width

	.StartUpPosition = 0    'allow manual positioning

	hwnd = FindWindow("ThunderDFrame", .Caption)
	CentreUserForm hwnd
    End With
End Sub

Public Sub UserForm_Show(ByVal ErrNum As Long, ByVal ErrMod As String, ByVal ErrDesc As String)
    With ErrorForm
	.ErrorNum.Value = "Run-time error '" & ErrNum & "':"
	.ErrorDesc.Value = ErrDesc
	.Caption = ErrMod & " Error"
    End With

    ErrorForm.Show
End Sub
