VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIndent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'              Converter Indent Class
'
' Note: Class is PublicNotCreatable
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private IndentNum As Long       'indent by this number of spaces

Private OptionOutputTypeNames As Boolean        'include type name in output
Private OptionWrapStringsInQuotes As Boolean    'wrap strings in quotes
Private OptionAllowEvents As Boolean            'allow the use of DoEvents

'######################## Indent Methods ##########################

Private Sub Class_Initialize()
    IndentNum = 4     'default

    OptionOutputTypeNames = True
    OptionWrapStringsInQuotes = True
    OptionAllowEvents = True
End Sub

'***************************************
'
'           Debug.print Object
'
'***************************************

Public Sub DebugPrint(ByVal Result As Object)
    Debug.Print Format(Result)
End Sub

'***************************************
'
'        Format Object to a File
'
'***************************************

Public Function FormatFile(Result As Object, ByVal filePath As String) As Boolean
    Dim FSO As FileSystemObject
    Dim TS As TextStream

    Set FSO = New FileSystemObject

    If Not FSO.FileExists(filePath) Then
        Set TS = FSO.CreateTextFile(filePath)
    Else
        Set TS = FSO.OpenTextFile(filePath, ForAppending, TristateTrue)
    End If

    TS.WriteLine ""
    TS.WriteLine Format(Result)

    TS.Close

    FormatFile = True

    Set FSO = Nothing
    Set TS = Nothing
End Function

'***************************************
'
'        Format Object as String
'
'***************************************

Public Function Format(Result As Object) As String
    Dim Buffer As New CBuffer

    Format = ""
    FormatEx Result, 0, Buffer
    Format = Buffer.Value
End Function

Private Sub FormatEx(ByVal Result As Object, ByVal Idx As Long, ByRef Buffer As CBuffer)
    Dim Key As Variant
    Dim keyType As String
    Dim i As Long

    keyType = TypeName(Result)
    If keyType = "Dictionary" Then
        For Each Key In Result.Keys
            keyType = TypeName(Result(Key))

            If keyType = "Dictionary" Then
                Buffer.Append Space(Idx) & Key & ": " & vbLf
                FormatEx Result(Key), Idx + IndentNum, Buffer
            ElseIf keyType = "Collection" Then
                Buffer.Append Space(Idx) & Key & ": " & vbLf
                FormatEx Result(Key), Idx + IndentNum, Buffer
            Else
                Buffer.Append Space(Idx) & Key & IIf(OptionOutputTypeNames, " (" & keyType & "): ", ": ") & JsonValue(Result(Key)) & vbLf
            End If

            If OptionAllowEvents Then DoEvents
        Next Key
    ElseIf keyType = "Collection" Then
        For i = 1 To Result.Count
            keyType = TypeName(Result(i))

            If keyType = "Dictionary" Then
                Buffer.Append Space(Idx) & i & ": " & vbLf
                FormatEx Result(i), Idx + IndentNum, Buffer
            ElseIf keyType = "Collection" Then
                Buffer.Append Space(Idx) & i & ": " & vbLf
                FormatEx Result(i), Idx + IndentNum, Buffer
            Else
                Buffer.Append Space(Idx) & i & IIf(OptionOutputTypeNames, " (" & keyType & "): ", ": ") & JsonValue(Result(i)) & vbLf
            End If

            If OptionAllowEvents Then DoEvents
        Next i
    End If
End Sub

Private Function JsonValue(ByVal Val As Variant) As String
    Select Case VarType(Val)
        Case vbString
            JsonValue = """" & Val & """"         'add quotes so you can see empty strings
        Case vbBoolean
            JsonValue = LCase$(CStr(Val))          'match json lowercase standard
        Case vbNull
            JsonValue = "null"
        Case Else
            If IsEmpty(Val) Then
                JsonValue = "undefined"
            Else
                JsonValue = CStr(Val)
            End If
    End Select
End Function

'################### Get/Let Button Properties ####################

'OptionAllowEvents
Public Property Get AllowEvents() As Boolean
    AllowEvents = OptionAllowEvents
End Property

Public Property Let AllowEvents(Value As Boolean)
    OptionAllowEvents = Value
End Property

'OutputTypeNames
Public Property Get OutputTypeNames() As Boolean
    OutputTypeNames = OptionOutputTypeNames
End Property

Public Property Let OutputTypeNames(Value As Boolean)
    OptionOutputTypeNames = Value
End Property

'WrapStringsInQuotes
Public Property Get WrapStringsInQuotes() As Boolean
    WrapStringsInQuotes = OptionWrapStringsInQuotes
End Property

Public Property Let WrapStringsInQuotes(Value As Boolean)
    OptionWrapStringsInQuotes = Value
End Property
