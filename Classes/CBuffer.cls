VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------
'
'              Converter Buffer Class
'
' Note: Class is Private
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
' Based on the work by Tim Hall from the site:
'   https://github.com/VBA-tools/VBA-JSON
'
'---------------------------------------------------

Private Buf As String       'buffer
Private BufPos As Long      'current position in buffer
Private BufLen As Long      'length of buffer

'######################## Buffer Methods ##########################

'***************************************
'
'        Append a Str to Buffer
'
'***************************************

Public Sub Append(Str As Variant)
    Dim AppendLen As Long, RequiredLen As Long, AddedLen As Long

    AppendLen = Len(Str)
    RequiredLen = AppendLen + BufPos

    If RequiredLen > BufLen Then
	AddedLen = IIf(AppendLen > BufLen, AppendLen, BufLen)   'at least double buffer space

	Buf = Buf & Space(AddedLen)
	BufLen = BufLen + AddedLen
    End If

    Mid$(Buf, BufPos + 1, AppendLen) = CStr(Str)
    BufPos = BufPos + AppendLen
End Sub

'################### Get/Let Button Properties ####################

'Value
Public Property Get Value() As Variant
    If BufPos > 0 Then
	Value = Left$(Buf, BufPos)
    End If
End Property

Public Property Let Value(Str As Variant)
    BufPos = 0  'keep exiting buffer space
    If Len(Str) > 0 Then
	Append Str
    End If
End Property
