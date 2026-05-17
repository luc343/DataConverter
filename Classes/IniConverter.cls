VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'                IniConverter Class
'
' Note: Class is PublicNotCreatable

' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private IniStr As String    'ini string
Private IniLen As Long      'ini string length
Private Idx As Long         'index into Ini string

Private LineNum As Long     'error line number
Private LineCol As Long     'error colunm

Private OptionUseDoubleForLargeNumbers As Boolean   'false = use strings, true = use doubles
Private OptionUseTypesForValues As Boolean          'determine value variable types
Private OptionAllowEvents As Boolean                'allow the use of DoEvents

Private ErrorMod As String      'error module
Private ErrorDesc As String     'error discription
Private ErrorNum As Long        'error number

Private Const CHAR_CR As Long = 8592        'left pointing arrow
Private Const CHAR_LF As Long = 8595        'down pointing arrow

'######################## INI Methods ##########################

Private Sub Class_Initialize()
    OptionAllowEvents = True
End Sub

'***************************************
'
'          Parse an Ini String
'
'***************************************

Public Function Parse(ByVal IniString As String) As Object
    Dim Properties As New Scripting.Dictionary
    Dim ErrIni As String
    Dim SecName As Variant
    Dim Key As Variant

    IniConverterErr = False

    Idx = 1
    IniStr = IniString
    IniLen = Len(IniStr)

    LineNum = 1
    LineCol = 1

    SkipBlanks

    Set Parse = New Scripting.Dictionary

    On Error GoTo Err
    While Idx <= IniLen
        If Mid$(IniStr, Idx, 1) <> "[" Then
            SecName = "Global"
            SkipSpaces
        Else
            SecName = ParseSection()
        End If
        Set Properties = ParseProperties()

        If Parse.Exists(SecName) Then
            For Each Key In Properties.Keys
                Parse.Item(SecName)(Key) = Properties(Key)      'will overwrite an exist key
            Next Key
        Else
            Parse.Add SecName, Properties
        End If
        SkipBlanks
    Wend
    On Error GoTo 0
    Exit Function

Err:
    ErrorNum = Err.Number
    ErrorMod = Err.Source
    ErrorDesc = Err.Description

    On Error GoTo 0
    ErrorForm.UserForm_Show ErrorNum, ErrorMod, ErrorDesc

    Set Parse = Nothing
    IniConverterErr = True
End Function

'***************************************
'
'     Parse an Ini String from File
'
'***************************************

Public Function ParseFile(ByVal filePath As String) As Object
    Dim IniString As String, ErrorIni As String

    IniConverterErr = False

    On Error GoTo Err
    If Not FileExists(filePath) Then
        Err.Raise 150, "IniConverter", "Path '" & filePath & "' not found!"
    End If
    On Error GoTo 0

    IniString = ReadTextFile(filePath)
    Set ParseFile = IniConverter.Parse(IniString)
    Exit Function

Err:
    ErrorNum = Err.Number
    ErrorMod = Err.Source
    ErrorDesc = Err.Description

    On Error GoTo 0
    ErrorForm.UserForm_Show ErrorNum, ErrorMod, ErrorDesc
    Set ParseFile = Nothing

    IniConverterErr = True
End Function

Private Function ParseSection() As String
    Dim Ch As String
    Dim InName As Boolean
    Dim pos As Long

    ParseSection = ""
    Do While Idx > 0 And Idx <= IniLen
        Ch = Mid$(IniStr, Idx, 1)

        Select Case Ch
        Case ";", "#"
            If InName Then
                Err.Raise 151, "IniConverter", ErrorStr("Expecting ']'")
            End If

            Idx = Idx + 1
            SkipComment
        Case vbLf, vbCr
            If ParseSection = "" Then
                Err.Raise 152, "IniConverter", ErrorStr("Expecting '[' or ']'")
            End If

            SkipNewline
            Exit Do
        Case "["
            InName = True
            Idx = Idx + 1
            pos = Idx
        Case "]"
            If InName Then
                If Idx - pos = 0 Then
                    Err.Raise 153, "IniConverter", ErrorStr("Expecting 'CHARACTOR'")
                End If
                ParseSection = Mid$(IniStr, pos, Idx - pos)
                InName = False
                Idx = Idx + 1
                SkipSpaces
            Else
                Err.Raise 154, "IniConverter", ErrorStr("Expecting '['")
            End If
        Case Else
            If InName Then
                Idx = Idx + 1
            Else
                Err.Raise 155, "IniConverter", ErrorStr("Expecting 'CR', 'LF', ':' or '#'")
            End If
        End Select
    Loop

    If Idx > IniLen And ParseSection = "" Then
        Err.Raise 156, "IniConverter", ErrorStr("Expecting ']'")
    End If
End Function

Private Function ParseProperties() As Scripting.Dictionary
    Dim Ch As String

    Set ParseProperties = New Scripting.Dictionary

    SkipBlanks

    Do While Idx > 0 And Idx <= IniLen
        Ch = Mid$(IniStr, Idx, 1)

        If Ch = "[" Then
            Exit Do
        Else
            ParseProperty ParseProperties
            SkipBlanks
        End If
    Loop
End Function

Private Function ParseProperty(ByRef Properties As Scripting.Dictionary)
    Dim InKey As Boolean
    Dim posKey As Long, posVal As Long
    Dim Ch As String, Key As String, Value As String

    InKey = True
    posKey = Idx
    posVal = 0
    Do While Idx > 0 And Idx <= IniLen
        Ch = Mid$(IniStr, Idx, 1)

        Select Case Ch
        Case "["
            Exit Do
        Case ";", "#"
            Idx = Idx + 1
            SkipComment
        Case vbLf, vbCr
            If InKey Then   'there was no =
                Key = Mid$(IniStr, posKey, Idx - posKey)
                If Key = "" Then
                    Err.Raise 157, "IniConverter", ErrorStr("Expecting 'CHARACTER'")
                End If
            End If

            Value = ""
            If posVal <> 0 Then
                Value = Mid$(IniStr, posVal, Idx - posVal)
            End If
            Properties.Item(Trim$(Key)) = Value

            SkipBlanks
            InKey = True
            posKey = Idx
            posVal = 0
        Case "="
            If InKey Then
                Key = Mid$(IniStr, posKey, Idx - posKey)
                If Key = "" Then
                    Err.Raise 158, "IniConverter", ErrorStr("Expecting 'CHARACTER'")
                End If

                InKey = False
                Idx = Idx + 1
                posVal = Idx
            Else
                Idx = Idx + 1
            End If
        Case Else
            Idx = Idx + 1
        End Select
    Loop

    If Idx > IniLen And posKey <= IniLen Then       'line did not end in CR or LF
        If InKey Then
            Key = Mid$(IniStr, posKey, Idx - posKey)
            If Key = "" Then
                Err.Raise 159, "IniConverter", ErrorStr("Expecting 'CHARACTER'")
            End If
        End If

        Value = ""
        If posVal <> 0 Then
            Value = Mid$(IniStr, posVal, Idx - posVal)
        End If
        Properties.Item(Trim$(Key)) = Value
    End If
End Function

Private Function IsSpace(Ch As String) As Boolean
    IsSpace = (Ch = " " Or Ch = vbTab)
End Function

Private Sub SkipSpaces()
    Do While Idx > 0 And Idx <= IniLen And IsSpace(Mid$(IniStr, Idx, 1))
        Idx = Idx + 1
    Loop

    If OptionAllowEvents Then DoEvents
End Sub

Private Function IsNewline(Ch As String) As Boolean
    IsNewline = (Ch = vbLf Or Ch = vbCr)
    If Ch = vbLf Then       'covers both CRLF and LF
        LineNum = LineNum + 1
        LineCol = Idx + 1
    End If
End Function

Private Sub SkipNewline()
    Do While Idx > 0 And Idx <= IniLen And IsNewline(Mid$(IniStr, Idx, 1))
        Idx = Idx + 1
    Loop

    If OptionAllowEvents Then DoEvents
End Sub

Private Sub SkipComment()
    Do While Idx > 0 And Idx <= IniLen And Not IsNewline(Mid$(IniStr, Idx, 1))
        Idx = Idx + 1
    Loop

    If OptionAllowEvents Then DoEvents
End Sub

Private Sub SkipBlanks()
    Dim Ch As String
    Dim Done As Boolean

    Done = False
    Do While Idx > 0 And Idx <= IniLen And Not Done
        SkipSpaces

        Ch = Mid$(IniStr, Idx, 1)
        If Ch = ";" Or Ch = "#" Then
            Idx = Idx + 1
            SkipComment
        End If
        If IsNewline(Mid$(IniStr, Idx, 1)) Then
            SkipNewline
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function ErrorStr(ErrorMessage As String)
    Dim StartIndex As Long, StopIndex As Long
    Dim Out As String

    StartIndex = Idx - 50
    StopIndex = Idx + 50
    If StartIndex <= 0 Then
        StartIndex = 1
    End If
    If StopIndex > IniLen Then
        StopIndex = IniLen
    End If

    Out = Mid$(IniStr, StartIndex, StopIndex - StartIndex + 1)
    Out = Replace(Replace(Replace(Out, vbCr, ChrW(CHAR_CR)), vbLf, ChrW(CHAR_LF)), vbTab, " ")

    'Example:
    'Error parsing Ini:
    '[abc]     a
    '          ^
    'Expecting 'CR' or 'LF' at index 11

    ErrorStr = "Error parsing Ini:" & vbLf & _
                         Out & vbLf & _
                         Space(Idx - StartIndex) & "^" & vbLf & _
                         ErrorMessage & " at line " & LineNum & ", colunm " & Idx - LineCol + 1 & ", position " & Idx
End Function

'################### Get/Let Properties ####################

'ErrNum
Public Property Get ErrNum() As Long
    ErrNum = ErrorNum
End Property

'ErrDesc
Public Property Get ErrDesc() As String
    ErrDesc = ErrorDesc
End Property

'ErrMod
Public Property Get ErrMod() As String
    ErrMod = ErrorMod
End Property

'OptionAllowEvents
Public Property Get AllowEvents() As Boolean
    AllowEvents = OptionAllowEvents
End Property

Public Property Let AllowEvents(Value As Boolean)
    OptionAllowEvents = Value
End Property

'UseDoubleForLargeNumbers
Public Property Get UseDoubleForLargeNumbers() As Boolean
    UseDoubleForLargeNumbers = OptionUseDoubleForLargeNumbers
End Property

Public Property Let UseDoubleForLargeNumbers(Value As Boolean)
    OptionUseDoubleForLargeNumbers = Value
End Property

'UseTypesForValues
Public Property Get UseTypesForValues() As Boolean
    UseTypesForValues = OptionUseTypesForValues
End Property

Public Property Let UseTypesForValues(Value As Boolean)
    OptionUseTypesForValues = Value
End Property
