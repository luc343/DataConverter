VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'               JsonConverter Class
'
' Note: Class is PublicNotCreatable
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

'Note: Use ReturnJsonErrors for JsonConverter to return all errors as a valid JSON Object. Does not raise errors

Private JsonStr As String   'json string
Private JsonLen As Long     'json string length
Private Idx As Long         'index into json string

Private LineNum As Long     'error line number
Private LineCol As Long     'error colunm

'available options
Private OptionUseDoubleForLargeNumbers As Boolean   'false = use strings, true = use doubles
Private OptionAllowUnquotedKeys As Boolean          'false = use quoted,  true = use " or '
Private OptionEscapeSolidus As Boolean              'false = use /,       true = use \/
Private OptionReturnJsonErrors As Boolean           'return errors in json
Private OptionAllowEvents As Boolean                'allow the use of DoEvents

Private ErrorMod As String      'error module
Private ErrorDesc As String     'error discription
Private ErrorNum As Long        'error number

Private Const CHAR_CR As Long = 8592        'left pointing arrow
Private Const CHAR_LF As Long = 8595        'down pointing arrow

'######################## Json Methods ##########################

Private Sub Class_Initialize()
    OptionAllowEvents = True
End Sub

'***************************************
'
'        Parse a Json String
'
'***************************************

Public Function Parse(ByVal JsonString As String) As Object
    Dim ErrorJson As String

    JsonConverterErr = False

    Idx = 1
    JsonStr = JsonString
    JsonLen = Len(JsonStr)

    LineNum = 1
    LineCol = 1

    SkipSpaces

    On Error GoTo Err
    Select Case Mid$(JsonStr, Idx, 1)
        Case "{"
            Set Parse = ParseObject()    'dictionary
        Case "["
            Set Parse = ParseArray()     'collection
        Case Else
            Err.Raise 101, "JSONConverter", ErrorStr("Expecting '{' or '['")
    End Select
    On Error GoTo 0
    Exit Function

Err:
    ErrorNum = Err.Number
    ErrorMod = Err.Source
    ErrorDesc = Err.Description

    On Error GoTo 0
    If OptionReturnJsonErrors Then          'return json error objects
        ErrorJson = ErrorJSONStr(ErrorNum, ErrorMod, ErrorDesc)
        Set Parse = Parse(ErrorJson)        'could cause an infinite loop
    Else
        ErrorForm.UserForm_Show ErrorNum, ErrorMod, ErrorDesc
        Set Parse = Nothing
    End If

    JsonConverterErr = True
End Function

'***************************************
'
'     Parse a Json String from File
'
'***************************************

Public Function ParseFile(ByVal filePath As String) As Object
    Dim JsonString As String, ErrorJson As String

    JsonConverterErr = False

    On Error GoTo Err
    If Not FileExists(filePath) Then
        Err.Raise 108, "JSONConverter", "Path '" & filePath & "' not found!"
    End If
    On Error GoTo 0

    JsonString = ReadTextFile(filePath)
    Set ParseFile = JsonConverter.Parse(JsonString)
    Exit Function

Err:
    ErrorNum = Err.Number
    ErrorMod = Err.Source
    ErrorDesc = Err.Description

    On Error GoTo 0
    If OptionReturnJsonErrors Then              'return json error objects
        ErrorJson = ErrorJSONStr(ErrorNum, ErrorMod, ErrorDesc)
        Set ParseFile = Parse(ErrorJson)        'could cause an infinite loop
    Else
        ErrorForm.UserForm_Show ErrorNum, ErrorMod, ErrorDesc
        Set ParseFile = Nothing
    End If

    JsonConverterErr = True
End Function

'***************************************
'
'       Convert to a Json String
'
'***************************************

Public Function ToString(ByVal JsonObj As Variant, Optional ByVal WhiteSpace As Variant, Optional ByVal CurIndent As Long = 0) As String
    Dim Index As Long, LBnd As Long, UBnd As Long, LBnd2D As Long, UBnd2D As Long, Index2D As Long
    Dim DateStr As String, Converted As String, Indent As String, InnerIndent As String
    Dim IsFirstItem As Boolean, IsFirstItem2D As Boolean, SkipItem As Boolean, IsPrettyPrint As Boolean
    Dim Key As Variant, Value As Variant
    Dim Buffer As New CBuffer

    LBnd = -1
    UBnd = -1
    LBnd2D = -1
    UBnd2D = -1
    IsFirstItem = True
    IsFirstItem2D = True
    IsPrettyPrint = Not IsMissing(WhiteSpace)

    Select Case VarType(JsonObj)
    Case vbNull
        ToString = "null"
    Case vbDate     'date
        DateStr = ConvertToIso(CDate(JsonObj))

        ToString = """" & DateStr & """"
    Case vbString   'string (or large number encoded as string)
        If Not OptionUseDoubleForLargeNumbers And StringIsLargeNumber(JsonObj) Then
            ToString = JsonObj
        Else
            ToString = """" & Encode(JsonObj) & """"
        End If
    Case vbBoolean
        If JsonObj Then
            ToString = "true"
        Else
            ToString = "false"
        End If
    Case vbArray To vbArray + vbByte
        If IsPrettyPrint Then
            If VarType(WhiteSpace) = vbString Then
                Indent = String(CurIndent + 1, WhiteSpace)
                InnerIndent = String(CurIndent + 2, WhiteSpace)
            Else
                Indent = Space$((CurIndent + 1) * WhiteSpace)
                InnerIndent = Space$((CurIndent + 2) * WhiteSpace)
            End If
        End If


        Buffer.Append "["   'array

        On Error Resume Next

        LBnd = LBound(JsonObj, 1)
        UBnd = UBound(JsonObj, 1)
        LBnd2D = LBound(JsonObj, 2)
        UBnd2D = UBound(JsonObj, 2)

        If LBnd >= 0 And UBnd >= 0 Then
            For Index = LBnd To UBnd
                If IsFirstItem Then
                    IsFirstItem = False
                Else

                    Buffer.Append ","   'append comma to previous line
                End If

                If LBnd2D >= 0 And UBnd2D >= 0 Then     ' 2d array
                    If IsPrettyPrint Then
                        Buffer.Append vbLf
                    End If
                    Buffer.Append Indent & "["

                    For Index2D = LBnd2D To UBnd2D
                        If IsFirstItem2D Then
                            IsFirstItem2D = False
                        Else
                            Buffer.Append ","
                        End If

                        Converted = ToString(JsonObj(Index, Index2D), WhiteSpace, CurIndent + 2)
                        If Converted = "" Then                              'for arrays/collections, undefined (empty/nothing) is treated as null
                            If IsUndefined(JsonObj(Index, Index2D)) Then    '(nest to only check if converted = "")
                                Converted = "null"
                            End If
                        End If

                        If IsPrettyPrint Then
                            Converted = vbLf & InnerIndent & Converted
                        End If

                        Buffer.Append Converted
                    Next Index2D

                    If IsPrettyPrint Then
                        Buffer.Append vbLf
                    End If

                    Buffer.Append Indent & "]"
                    IsFirstItem2D = True
                Else                                    '1d array
                    Converted = ToString(JsonObj(Index), WhiteSpace, CurIndent + 1)
                    If Converted = "" Then                                  'for arrays/collections, undefined (empty/nothing) is treated as null
                        If IsUndefined(JsonObj(Index)) Then                 '(nest to only check if converted = "")
                            Converted = "null"
                        End If
                    End If

                    If IsPrettyPrint Then
                        Converted = vbLf & Indent & Converted
                    End If

                    Buffer.Append Converted
                End If
            Next Index
        End If

        On Error GoTo 0

        If IsPrettyPrint Then
            Buffer.Append vbLf

            If VarType(WhiteSpace) = vbString Then
                Indent = String(CurIndent, WhiteSpace)
            Else
                Indent = Space(CurIndent * WhiteSpace)
            End If
        End If

        Buffer.Append Indent & "]"
        ToString = Buffer.Value

    Case vbObject   'dictionary or collection
        If IsPrettyPrint Then
            If VarType(WhiteSpace) = vbString Then
                Indent = String(CurIndent + 1, WhiteSpace)
            Else
                Indent = Space$((CurIndent + 1) * WhiteSpace)
            End If
        End If

        If TypeName(JsonObj) = "Dictionary" Then        'dictionary
            Buffer.Append "{"
            For Each Key In JsonObj.Keys
                Converted = ToString(JsonObj(Key), WhiteSpace, CurIndent + 1)       'for objects, undefined (Empty/Nothing) is not added to object
                If Converted = "" Then                                              '(nest to only check if converted = "")
                    SkipItem = IsUndefined(JsonObj(Key))
                Else
                    SkipItem = False
                End If

                If Not SkipItem Then
                    If IsFirstItem Then
                        IsFirstItem = False
                    Else
                        Buffer.Append ","
                    End If

                    If IsPrettyPrint Then
                        Converted = vbLf & Indent & """" & Key & """: " & Converted
                    Else
                        Converted = """" & Key & """:" & Converted
                    End If

                    Buffer.Append Converted
                End If
            Next Key

            If IsPrettyPrint Then
                Buffer.Append vbLf

                If VarType(WhiteSpace) = vbString Then
                    Indent = String(CurIndent, WhiteSpace)
                Else
                    Indent = Space(CurIndent * WhiteSpace)
                End If
            End If

            Buffer.Append Indent & "}"

        ElseIf TypeName(JsonObj) = "Collection" Then    'collection
            Buffer.Append "["
            For Each Value In JsonObj
                If IsFirstItem Then
                    IsFirstItem = False
                Else
                    Buffer.Append ","
                End If
                Converted = ToString(Value, WhiteSpace, CurIndent + 1)
                If Converted = "" Then                                              'for arrays/collections, undefined (empty/nothing) is treated as null
                    If IsUndefined(Value) Then                                      '(nest to only check if converted = "")
                        Converted = "null"
                    End If
                End If

                If IsPrettyPrint Then
                    Converted = vbLf & Indent & Converted
                End If

                Buffer.Append Converted
            Next Value

            If IsPrettyPrint Then
                Buffer.Append vbLf

                If VarType(WhiteSpace) = vbString Then
                    Indent = String(CurIndent, WhiteSpace)
                Else
                    Indent = Space(CurIndent * WhiteSpace)
                End If
            End If

            Buffer.Append Indent & "]"
        End If

        ToString = Buffer.Value
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal   'number (use decimals for numbers)
        ToString = Replace(JsonObj, ",", ".")
    Case Else       'vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        On Error Resume Next
        ToString = JsonObj      'use VBA's built-in to-string
        On Error GoTo 0
    End Select
End Function

Private Function ParseObject() As Scripting.Dictionary
    Dim Key As String
    Dim NextChar As String

    Set ParseObject = New Scripting.Dictionary

    SkipSpaces
    If Mid$(JsonStr, Idx, 1) <> "{" Then
        Err.Raise 102, "JSONConverter", ErrorStr("Expecting '{'")
    End If

    Idx = Idx + 1

    Do
        SkipSpaces

        If Mid$(JsonStr, Idx, 1) = "}" Then
            Idx = Idx + 1
            Exit Function
        ElseIf Mid$(JsonStr, Idx, 1) = "," Then
            Idx = Idx + 1
            SkipSpaces
        End If

        Key = ParseKey()
        NextChar = Peek()
        If NextChar = "[" Or NextChar = "{" Then
            Set ParseObject.Item(Key) = ParseValue()
        Else
            ParseObject.Item(Key) = ParseValue()
        End If
    Loop
End Function

Private Function ParseArray() As Collection
    Set ParseArray = New Collection

    SkipSpaces
    If Mid$(JsonStr, Idx, 1) <> "[" Then
        Err.Raise 103, "JSONConverter", ErrorStr("Expecting '['")
    End If

    Idx = Idx + 1

    Do
        SkipSpaces
        If Mid$(JsonStr, Idx, 1) = "]" Then
            Idx = Idx + 1
            Exit Function
        ElseIf Mid$(JsonStr, Idx, 1) = "," Then
            Idx = Idx + 1
            SkipSpaces
        End If

        ParseArray.Add ParseValue()
    Loop
End Function

Private Function ParseValue() As Variant
    SkipSpaces
    Select Case Mid$(JsonStr, Idx, 1)
    Case "{"
        Set ParseValue = ParseObject()
    Case "["
        Set ParseValue = ParseArray()
    Case """", "'"
        ParseValue = ParseString()
    Case Else
        If Mid$(JsonStr, Idx, 4) = "true" Then
            ParseValue = True
            Idx = Idx + 4
        ElseIf Mid$(JsonStr, Idx, 5) = "false" Then
            ParseValue = False
            Idx = Idx + 5
        ElseIf Mid$(JsonStr, Idx, 4) = "null" Then
            ParseValue = Null
            Idx = Idx + 4
        ElseIf InStr("+-0123456789", Mid$(JsonStr, Idx, 1)) Then
            ParseValue = ParseNumber()
        Else
            Err.Raise 104, "JSONConverter", ErrorStr("Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function ParseString() As String
    Dim Quote As String, Ch As String, Code As String
    Dim Buffer As New CBuffer

    SkipSpaces

    Quote = Mid$(JsonStr, Idx, 1)
    Idx = Idx + 1

    Do While Idx > 0 And Idx <= JsonLen
        Ch = Mid$(JsonStr, Idx, 1)

        Select Case Ch
        Case "\"        'escaped string, \\, or \/
            Idx = Idx + 1
            Ch = Mid$(JsonStr, Idx, 1)

            Select Case Ch
            Case """", "\", "/", "'"
                Buffer.Append Ch
                Idx = Idx + 1
            Case "b"
                Buffer.Append vbBack
                Idx = Idx + 1
            Case "f"
                Buffer.Append vbFormFeed
                Idx = Idx + 1
            Case "n"
                Buffer.Append vbCrLf
                Idx = Idx + 1
            Case "r"
                Buffer.Append vbCr
                Idx = Idx + 1
            Case "t"
                Buffer.Append vbTab
                Idx = Idx + 1
            Case "u"            'unicode character escape (e.g. \u00a9 = Copyright)
                Idx = Idx + 1
                Code = Mid$(JsonStr, Idx, 4)
                Buffer.Append ChrW(Val("&h" + Code))
                Idx = Idx + 4
            End Select
        Case Quote
            ParseString = Buffer.Value
            Idx = Idx + 1
            Exit Function
        Case Else
            Buffer.Append Ch
            Idx = Idx + 1
        End Select
    Loop
End Function

Private Function ParseNumber() As Variant
    Dim Ch As String, Value As String
    Dim IsLargeNumber As Boolean

    SkipSpaces

    While Idx > 0 And Idx <= JsonLen
        Ch = Mid$(JsonStr, Idx, 1)

        If InStr("+-0123456789.eE", Ch) Then        'unlikely to have massive number, so use simple append rather than buffer here
            Value = Value & Ch
            Idx = Idx + 1
        Else
            IsLargeNumber = IIf(InStr(Value, "."), Len(Value) >= 17, Len(Value) >= 16)

            If Not OptionUseDoubleForLargeNumbers And IsLargeNumber Then
                ParseNumber = Value
            Else
                ParseNumber = Val(Value)        'Val() does not use regional settings, so guard for comma is not needed
            End If
            Exit Function
        End If
    Wend
End Function

Private Function ParseKey() As String
    Dim Ch As String

    If Mid$(JsonStr, Idx, 1) = """" Or Mid$(JsonStr, Idx, 1) = "'" Then
        ParseKey = ParseString()
    ElseIf OptionAllowUnquotedKeys Then
        Do While Idx > 0 And Idx <= JsonLen
            Ch = Mid$(JsonStr, Idx, 1)
            If (Ch <> " ") And (Ch <> ":") Then
                ParseKey = ParseKey & Ch
                Idx = Idx + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 105, "JSONConverter", ErrorStr("Expecting '""' or '''")
    End If


    SkipSpaces
    If Mid$(JsonStr, Idx, 1) <> ":" Then    'check for colon and skip if present or throw if not present
        Err.Raise 106, "JSONConverter", ErrorStr("Expecting ':'")
    Else
        Idx = Idx + 1
    End If
End Function

Private Function IsUndefined(ByVal Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbEmpty
            IsUndefined = True
        Case vbObject
            If TypeName(Value) = "Empty" Or TypeName(Value) = "Nothing" Then
                IsUndefined = True
            End If
        End Select
End Function

Private Function Encode(ByVal Text As Variant) As String
    Dim AscCode As Long, i As Long
    Dim Buffer As New CBuffer
    Dim Ch As String

    For i = 1 To Len(Text)
        Ch = Mid$(Text, i, 1)
        AscCode = AscW(Ch)

        If AscCode < 0 Then     'fix unsigned long
            AscCode = AscCode + 65536
        End If

        Select Case AscCode
            Case 34
                Ch = "\"""
            Case 92
                Ch = "\\"
            Case 47
                If OptionEscapeSolidus Then
                    Ch = "\/"
                End If
            Case 8
                Ch = "\b"
            Case 12
                Ch = "\f"
            Case 10
                Ch = "\n"
            Case 13
                Ch = "\r"
            Case 9
                Ch = "\t"
            Case 0 To 31, 127 To 65535
                Ch = "\u" & Right$("0000" & Hex(AscCode), 4)
        End Select

        Buffer.Append Ch
    Next i

    Encode = Buffer.Value
End Function

Private Function Peek(Optional NumOfChars As Long = 1) As String
    SkipSpaces
    Peek = Mid$(JsonStr, Idx, NumOfChars)
End Function

Private Function IsSpace(Ch As String) As Boolean
    IsSpace = (Ch = " " Or Ch = vbLf Or Ch = vbCr Or Ch = vbTab)
    If Ch = vbLf Then       'covers both CRLF and LF
        LineNum = LineNum + 1
        LineCol = Idx + 1
    End If
End Function

Private Sub SkipSpaces()
    Do While Idx > 0 And Idx <= JsonLen And IsSpace(Mid$(JsonStr, Idx, 1))
        Idx = Idx + 1
    Loop

    If OptionAllowEvents Then DoEvents
End Sub

Private Function StringIsLargeNumber(Value As Variant) As Boolean
    Dim Length As Long, i As Long
    Dim ChCode As String

    Length = Len(Value)

    If Length >= 16 And Length <= 100 Then
        StringIsLargeNumber = True

        For i = 1 To Length
            ChCode = Asc(Mid$(Value, i, 1))
            Select Case ChCode
                Case 46, 48 To 57, 69, 101      'look for .|0-9|E|e
                    'fall through
                Case Else
                    StringIsLargeNumber = False
                    Exit Function
            End Select
        Next i
    End If
End Function

Private Function ConvertToIso(LocalDate As Date) As String
    On Error GoTo Err
    ConvertToIso = LocalDateToUTC(LocalDate)
    On Error GoTo 0
    Exit Function
Err:
    Err.Raise 107, "ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

Private Function ErrorStr(ErrorMessage As String)
    Dim StartIndex As Long, StopIndex As Long
    Dim Out As String

    StartIndex = Idx - 50
    StopIndex = Idx + 50
    If StartIndex <= 0 Then
        StartIndex = 1
    End If
    If StopIndex > JsonLen Then
        StopIndex = JsonLen
    End If

    Out = Mid$(JsonStr, StartIndex, StopIndex - StartIndex + 1)
    Out = Replace(Replace(Replace(Out, vbCr, ChrW(CHAR_CR)), vbLf, ChrW(CHAR_LF)), vbTab, " ")

    'Example:
    'Error parsing JSON:
    '{"abcde":True}
    '          ^
    'Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    ErrorStr = "Error parsing JSON:" & vbLf & _
                         Out & vbLf & _
                         Space(Idx - StartIndex) & "^" & vbLf & _
                         ErrorMessage & " at line " & LineNum & ", colunm " & Idx - LineCol + 1 & ", position " & Idx
End Function

Private Function ErrorJSONStr(ErrNum As Long, ModName As String, Desc As String) As String
    Dim Text As String, TextArr() As String
    Dim i As Long

    TextArr = Split(Desc, vbLf)

    ErrorJSONStr = "{""error_nr"":ERR_NR,""error_mod"":""ERR_MOD"",""error_details"":ARR_LINES}"

    Text = "["      'use array of lines
    For i = 0 To UBound(TextArr)
        Text = Text & """" & Replace(TextArr(i), """", "\""") & """"
        If i < UBound(TextArr) Then
            Text = Text & ","
        End If
    Next i

    Text = Text & "]"

    ErrorJSONStr = Replace(Replace(Replace(ErrorJSONStr, "ARR_LINES", Text), "ERR_MOD", ModName), "ERR_NR", ErrNum)
End Function

'################### Get/Let Properties ####################

Public Property Get ErrNum() As Long
    ErrNum = ErrorNum
End Property

Public Property Get ErrDesc() As String
    ErrDesc = ErrorDesc
End Property

Public Property Get ErrMod() As String
    ErrMod = ErrorMod
End Property

'UseDoubleForLargeNumbers
Public Property Get UseDoubleForLargeNumbers() As Boolean
    UseDoubleForLargeNumbers = OptionUseDoubleForLargeNumbers
End Property

Public Property Let UseDoubleForLargeNumbers(Value As Boolean)
    OptionUseDoubleForLargeNumbers = Value
End Property

'OptionAllowUnquotedKeys
Public Property Get AllowUnquotedKeys() As Boolean
    AllowUnquotedKeys = OptionAllowUnquotedKeys
End Property

Public Property Let AllowUnquotedKeys(Value As Boolean)
    OptionAllowUnquotedKeys = Value
End Property

'OptionEscapeSolidus
Public Property Get EscapeSolidus() As Boolean
    EscapeSolidus = OptionEscapeSolidus
End Property

Public Property Let EscapeSolidus(Value As Boolean)
    OptionEscapeSolidus = Value
End Property

'OptionReturnJsonErrors
Public Property Get ReturnJsonErrors() As Boolean
    ReturnJsonErrors = OptionReturnJsonErrors
End Property

Public Property Let ReturnJsonErrors(Value As Boolean)
    OptionReturnJsonErrors = Value
End Property

'OptionAllowEvents
Public Property Get AllowEvents() As Boolean
    AllowEvents = OptionAllowEvents
End Property

Public Property Let AllowEvents(Value As Boolean)
    OptionAllowEvents = Value
End Property
