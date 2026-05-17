Attribute VB_Name = "Lib_XML"
Option Explicit

'---------------------------------------------------
'
'                     Lib_XML
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Public XMLConverter As New XMLConverter     'xml parser

Public XMLConverterErr As Boolean

'***************************************
'
'      Escape Special Characters
'
'***************************************

Public Function EscapeXML(ByVal Text As String) As String
    Dim Out As String

    Out = Text

    Out = Replace(Out, "&", "&amp;")    'replace ampersand first (order matters)

    Out = Replace(Out, "<", "&lt;")     'replace the other 4 last
    Out = Replace(Out, ">", "&gt;")
    Out = Replace(Out, """", "&quot;")
    Out = Replace(Out, "'", "&apos;")

    EscapeXML = Out
End Function

'***************************************
'
'      Unescape Special Characters
'
'***************************************

Public Function UnescapeXML(ByVal Text As String) As String
    Dim Out As String

    Out = Text

    Out = Replace(Out, "&lt;", "<")     'replace the easy ones first
    Out = Replace(Out, "&gt;", ">")
    Out = Replace(Out, "&quot;", """")
    Out = Replace(Out, "&apos;", "'")

    Out = Replace(Out, "&amp;", "&")    'replace the ampersand last

    UnescapeXML = Out
End Function
