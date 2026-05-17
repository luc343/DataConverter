VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'                XMLConverter Class
'
' Copyright (c) Lucien Cinc 2025
'
' Note: Class is PublicNotCreatable
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
' Based on the work by Tim Hall from the site:
'   https://github.com/VBA-tools/VBA-XML
'
'---------------------------------------------------

Private XMLStr As String    'xml string
Private XMLLen As Long      'xml string length
Private Idx As Long         'index into xml string

Private LineNum As Long     'error line number
Private LineCol As Long     'error colunm

Private OptionExpandEntities As Boolean             'expand entities like '&amp;' to '&'
Private OptionAllowEvents As Boolean                'allow the use of DoEvents

Private ErrorMod As String      'error module
Private ErrorDesc As String     'error discription
Private ErrorNum As Long        'error number

Private Const TAG As String = "tag"                 'xml tag
Private Const ATTRIBS As String = "attribs"         'xml attributes
Private Const CONTENT As String = "content"         'xml content
Private Const ELEMENT As String = ""                'xml element

Private Const CHAR_CR As Long = 8592        'left pointing arrow
Private Const CHAR_LF As Long = 8595        'down pointing arrow

'######################## XML Methods ##########################

Private Sub Class_Initialize()
    OptionAllowEvents = True
End Sub


'***************************************
'
'          Parse an XML String
'
'***************************************

Public Function Parse(ByVal XMLString As String) As Object  'Scripting.Dictionary
    Dim ErrXML As String

    XMLConverterErr = False

    Set Parse = New Scripting.Dictionary

    Idx = 1
    XMLStr = XMLString
    XMLLen = Len(XMLStr)

    LineNum = 1
    LineCol = 1

    SkipSpaces

    On Error GoTo Err
    If Mid$(XMLStr, Idx, 1) <> "<" Then
        Err.Raise 120, "XMLConverter", ErrorStr("Expecting '<'")     'Error: Invalid XML string
    End If

    Set Parse = New Scripting.Dictionary

    Parse.Add "prolog", ParseProlog()           'only one PROLOG allowed
    Parse.Add "doctype", ParseDoctype()         'only one DOCTYPE allowed

    While Idx <= XMLLen
        Parse.Add ELEMENT & 1, ParseElement()
        SkipSpaces
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
    XMLConverterErr = True
End Function

'***************************************
'
'     Parse an XML String from File
'
'***************************************

Public Function ParseFile(ByVal filePath As String) As Object
    Dim XMLString As String, ErrorXML As String

    XMLConverterErr = False

    On Error GoTo Err
    If Not FileExists(filePath) Then
        Err.Raise 121, "XMLConverter", "Path '" & filePath & "' not found!"
    End If
    On Error GoTo 0

    XMLString = ReadTextFile(filePath)
    Set ParseFile = XMLConverter.Parse(XMLString)
    Exit Function

Err:
    ErrorNum = Err.Number
    ErrorMod = Err.Source
    ErrorDesc = Err.Description

    On Error GoTo 0
    ErrorForm.UserForm_Show ErrorNum, ErrorMod, ErrorDesc
    Set ParseFile = Nothing

    XMLConverterErr = True
End Function

'***************************************
'
'       Convert to a XML String
'
'***************************************

Public Function ToString(ByVal Dictionary As Scripting.Dictionary) As String
    'TODO
End Function

Private Function ParseProlog() As String
    Dim OpeningLevel As Long, pos As Long
    Dim Chars As String

    SkipSpaces

    If Mid$(XMLStr, Idx, 2) = "<?" Then
        pos = Idx
        Idx = Idx + 2

        Do
            Chars = Mid$(XMLStr, Idx, 2)

            If Idx + 1 > XMLLen Then
                Err.Raise 122, "XMLConverter", ErrorStr("Expecting '?>'")
            ElseIf OpeningLevel = 0 And Chars = "?>" Then
                Idx = Idx + 2
                Exit Do     'found matching closing tag, ?>
            ElseIf Chars = "<?" Then
                OpeningLevel = OpeningLevel + 1
                Idx = Idx + 2
            ElseIf Chars = "?>" Then
                OpeningLevel = OpeningLevel - 1
                Idx = Idx + 2
            Else
                Idx = Idx + 1
            End If
        Loop

        ParseProlog = Mid$(XMLStr, pos, Idx - pos)
    End If
End Function

Private Function ParseDoctype() As String
    Dim OpeningLevel As Long, pos As Long
    Dim Ch As String

    SkipSpaces
    If Mid$(XMLStr, Idx, 2) = "<!" Then
        pos = Idx
        Idx = Idx + 2

        Do
            Ch = Mid$(XMLStr, Idx, 1)
            Idx = Idx + 1

            If Idx > XMLLen Then
                Err.Raise 123, "XMLConverter", ErrorStr("Expecting '>'")
            ElseIf OpeningLevel = 0 And Ch = ">" Then
                Exit Do     'found matching closing tag, >
            ElseIf Ch = "<" Then
                OpeningLevel = OpeningLevel + 1
            ElseIf Ch = ">" Then
                OpeningLevel = OpeningLevel - 1
            End If
        Loop

        ParseDoctype = Mid$(XMLStr, pos, Idx - pos)
    End If
End Function

Private Function ParseElement() As Scripting.Dictionary
    Dim TagName As Variant
    Dim elnum As Long, pos As Long

    elnum = 1
    Set ParseElement = New Scripting.Dictionary

    SkipSpaces

    Do
        If Mid$(XMLStr, Idx, 1) <> "<" Then
            Err.Raise 124, "XMLConverter", ErrorStr("Expecting '<'")
        End If
        Idx = Idx + 1   'skip '<'

        If Idx > XMLLen Then
            Err.Raise 125, "XMLConverter", ErrorStr("Expecting '--', 'DOCTYPE', '[CDATA[' or 'ENTITY'")
        ElseIf Mid$(XMLStr, Idx, 1) = "!" Then
            If Mid$(XMLStr, Idx + 1, 2) = "--" Then             'skip comments
                pos = InStr(Idx + 3, XMLStr, "-->")
                If pos > 0 Then
                    Idx = pos + 3
                    SkipSpaces
                Else
                    Err.Raise 126, "XMLConverter", ErrorStr("Expecting '-->'")
                End If
            Else
                Err.Raise 127, "XMLConverter", ErrorStr("Expecting '--', 'DOCTYPE', '[CDATA[' or 'ENTITY'")
            End If
        Else
            Exit Do
        End If
    Loop

    TagName = ParseStartTag()                                       'parse start tag
    ParseElement.Add TAG, TagName

    ParseElement.Add ATTRIBS, ParseAttributes()                     'parse attributes

    If Mid$(XMLStr, Idx, 2) = "/>" Then
        Idx = Idx + 2       'skip over closing '/>'
        ParseElement.Add CONTENT, ""

    ElseIf Mid$(XMLStr, Idx, 1) = ">" Then
        Idx = Idx + 1       'skip over '>'
        ParseElement.Add CONTENT, ParseContent()                    'parse content

        Do
            If Idx > XMLLen Then
                Err.Raise 128, "XMLConverter", ErrorStr("Expecting '<' or '</'")
            ElseIf Mid$(XMLStr, Idx, 2) = "</" Then
                Idx = Idx + 2

                If TagName <> ParseEndTag() Then                        'parse end tag
                    Err.Raise 129, "XMLConverter", ErrorStr("Expecting '</" & TagName & ">'")
                End If

                If Mid$(XMLStr, Idx, 1) = ">" Then
                    Idx = Idx + 1
                Else
                    Err.Raise 130, "XMLConverter", ErrorStr("Expecting '>'")
                End If
                Exit Do
            ElseIf Mid$(XMLStr, Idx, 1) = "<" Then
                ParseElement.Add ELEMENT & elnum, ParseElement()        'parse element
                elnum = elnum + 1
                SkipSpaces

                'continue loop
            Else
                Err.Raise 131, "XMLConverter", ErrorStr("Expecting '<' or '</'")
            End If
        Loop
    Else
        Err.Raise 132, "XMLConverter", ErrorStr("Expecting '>' or '/>'")
    End If
Done:
End Function

Private Function ParseStartTag() As Variant
    Dim pos As Long
    Dim Ch As String

    pos = Idx
    Do
        Ch = Mid$(XMLStr, Idx, 1)

        Select Case Ch
        Case "<"
            Err.Raise 133, "XMLConverter", ErrorStr("Expecting ' ' or '>'")
        Case " ", ">"
            ParseStartTag = Mid$(XMLStr, pos, Idx - pos)

            SkipSpaces
            Exit Do
        Case Else
            If Not IsValidName(Ch, Idx - pos + 1) Then
                Err.Raise 134, "XMLConverter", ErrorStr("Invalid Character '" & Ch & "'")
            End If
            Idx = Idx + 1
        End Select

        If Idx > XMLLen Then
            Err.Raise 135, "XMLConverter", ErrorStr("Expecting ' ' or '>'")
        End If
    Loop
End Function

Private Function ParseEndTag() As Variant
    Dim pos As Long
    Dim Ch As String

    pos = Idx
    Do While Idx > 0 And Idx <= XMLLen
        Ch = Mid$(XMLStr, Idx, 1)

        Select Case Ch
        Case "<"
            Err.Raise 136, "XMLConverter", ErrorStr("Expecting ' ' or '>'")
        Case " ", ">", "/"
            ParseEndTag = Mid$(XMLStr, pos, Idx - pos)

            SkipSpaces
            Exit Do
        Case Else
            If Not IsValidName(Ch, Idx - pos + 1) Then
                Err.Raise 137, "XMLConverter", ErrorStr("Invalid Character '" & Ch & "'")
            End If
            Idx = Idx + 1
        End Select
    Loop
End Function

Private Function ParseAttributes() As Scripting.Dictionary
    Dim Ch As String, Quote As String, Name As String, Value As String
    Dim pos As Long

    Set ParseAttributes = New Scripting.Dictionary

    If Mid$(XMLStr, Idx, 2) = "/>" Then Exit Function

    SkipSpaces

    Quote = ""
    pos = Idx
    Do While Idx > 0 And Idx <= XMLLen
        Ch = Mid$(XMLStr, Idx, 1)

        Select Case Ch
        Case "="    'found end of attribute name
            Name = Mid$(XMLStr, pos, Idx - pos)   'extract name, skip =, reset start index, and check for quote

            Idx = Idx + 1

            Value = ParseString()
            ParseAttributes.Add Name, Value     'store name, value
        Case " ", ">", "/"
            If Ch = "/" And Mid$(XMLStr, Idx, 2) <> "/>" Then
                Idx = Idx + 1       'just a simple escape
            Else
                If Ch = ">" Or Ch = "/" Then
                    Exit Do
                Else
                    Idx = Idx + 1
                    SkipSpaces
                    pos = Idx
                End If
            End If
        Case Else
            Idx = Idx + 1
        End Select
    Loop
End Function

Private Function ParseContent() As Variant
    Dim Ch As String, unescbuf As String
    Dim isBuffering As Boolean
    Dim Buffer As New CBuffer

    SkipSpaces

    ParseContent = ""
    Do While Idx > 0 And Idx <= XMLLen
        Ch = Mid$(XMLStr, Idx, 1)

        Select Case Ch
        Case "<"
            If Not ParseCData(Buffer) Then
                ParseContent = Buffer.Value
                Exit Do
            End If
        Case Else
            If OptionExpandEntities Then
                ParseChar Ch, Buffer, unescbuf, isBuffering       'is slow
            Else
                Buffer.Append Ch
                Idx = Idx + 1
            End If
        End Select
    Loop
End Function

Private Function ParseCData(ByRef Buffer As CBuffer) As Boolean
    Dim Ch As String
    Dim Closed As Boolean

    ParseCData = (Mid$(XMLStr, Idx, 9) = "<![CDATA[")

    If ParseCData Then
        Closed = False
        Do While Idx > 0 And Idx <= XMLLen
            Ch = Mid$(XMLStr, Idx, 1)

            If Mid$(XMLStr, Idx, 3) = "]]>" Then
                Idx = Idx + 3
                Closed = True
                Exit Do
            Else
                Buffer.Append Ch
                Idx = Idx + 1
            End If
        Loop

        If Not Closed Then
            Err.Raise 138, "XMLConverter", ErrorStr("Expecting ']]>'")
        End If
    End If
End Function

Private Function ParseString() As String
    Dim Ch As String, unescbuf As String, Quote As String
    Dim isBuffering As Boolean
    Dim Buffer As New CBuffer
    Dim pos As Long

    Quote = Mid$(XMLStr, Idx, 1)
    If Quote <> """" And Quote <> "'" Then
        Err.Raise 139, "XMLConverter", ErrorStr("Expecting '""' or '''")
        Exit Function
    End If

    Idx = Idx + 1
    pos = Idx

    Do While Idx > 0 And Idx <= XMLLen
        Ch = Mid$(XMLStr, Idx, 1)

        Select Case Ch
        Case Quote
            ParseString = Buffer.Value
            Idx = Idx + 1   'skip quote
            Exit Do
        Case Else
            If OptionExpandEntities Then
                ParseChar Ch, Buffer, unescbuf, isBuffering       'is slower
            Else
                Buffer.Append Ch
                Idx = Idx + 1
            End If
        End Select
    Loop

    If Idx > XMLLen Then
        Err.Raise 141, "XMLConverter", ErrorStr("Expecting '""' or '''")
    End If
End Function

Private Sub ParseChar(Ch As String, ByRef Buffer As CBuffer, ByRef unescbuf As String, ByRef isBuffering As Boolean)
    If Ch = "&" Then
        If isBuffering Then         'hit a second '&' before a ';'
            Buffer.Append unescbuf  'should be an error - not a know entity
        End If

        unescbuf = Ch
        Idx = Idx + 1
        isBuffering = True

    ElseIf Ch = ";" And isBuffering Then
        unescbuf = unescbuf & Ch
        Idx = Idx + 1

        If unescbuf = "&amp;" Then     'order to unescape &amp; does not matter when not replacing
            Buffer.Append "&"
        ElseIf unescbuf = "&lt;" Then
            Buffer.Append "<"
        ElseIf unescbuf = "&gt;" Then
            Buffer.Append ">"
        ElseIf unescbuf = "&quot;" Then
            Buffer.Append """"
        ElseIf unescbuf = "&apos;" Then
            Buffer.Append "'"
        ElseIf Mid$(unescbuf, 1, 2) = "&#" Then
            If Mid$(unescbuf, 3, 1) = "x" Then
                Buffer.Append ChrW(CLng("&H" & Mid$(unescbuf, 4, Len(unescbuf) - 3 - 1)))     'hexadecimal
            Else
                Buffer.Append ChrW(CLng(Mid$(unescbuf, 3, Len(unescbuf) - 2 - 1)))            'decimal
            End If
        Else
            Buffer.Append unescbuf  'not a known entity
        End If

        unescbuf = ""
        isBuffering = False

    ElseIf isBuffering Then
        unescbuf = unescbuf & Ch
        Idx = Idx + 1

        If Len(unescbuf) > 10 Then     'it's to long
            Buffer.Append unescbuf     'should be an error - not a know entity

            unescbuf = ""
            isBuffering = False
        End If
    Else
        Buffer.Append Ch
        Idx = Idx + 1
    End If
End Sub

Private Function IsSpace(Ch As String) As Boolean
    IsSpace = (Ch = " " Or Ch = vbLf Or Ch = vbCr Or Ch = vbTab)
    If Ch = vbLf Then       'covers both CRLF and LF
        LineNum = LineNum + 1
        LineCol = Idx + 1
    End If
End Function

Private Sub SkipSpaces()
    Do While Idx > 0 And Idx <= XMLLen And IsSpace(Mid$(XMLStr, Idx, 1))
        Idx = Idx + 1
    Loop

    If OptionAllowEvents Then DoEvents
End Sub

Public Function IsValidName(ByRef Ch As String, ByVal pos As Long) As Boolean
    Select Case Asc(Ch)
        Case 65 To 90, 97 To 122, 95            'always valid is 'A'-'Z', 'a'-'z', '_'
            IsValidName = True
        Case 48 To 57, 45, 46
            If pos > 1 Then IsValidName = True  'valid when not first is '0'-'9', '-', '.'
    End Select
End Function

Private Function ErrorStr(ErrorMessage As String)
    Dim StartIndex As Long, StopIndex As Long
    Dim Out As String

    StartIndex = Idx - 50
    StopIndex = Idx + 50
    If StartIndex <= 0 Then
        StartIndex = 1
    End If
    If StopIndex > XMLLen Then
        StopIndex = XMLLen
    End If

    Out = Mid$(XMLStr, StartIndex, StopIndex - StartIndex + 1)
    Out = Replace(Replace(Replace(Out, vbCr, ChrW(CHAR_CR)), vbLf, ChrW(CHAR_LF)), vbTab, " ")

    'Example:
    'Error parsing XML:
    '<abc>1234</def>
    '          ^
    'Expecting '</abc>'

    ErrorStr = "Error parsing XML:" & vbLf & _
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

'OptionExpandEntities
Public Property Get ExpandEntities() As Boolean
    ExpandEntities = OptionExpandEntities
End Property

Public Property Let ExpandEntities(Value As Boolean)
    OptionExpandEntities = Value
End Property
