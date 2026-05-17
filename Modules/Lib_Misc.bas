Attribute VB_Name = "Lib_Misc"
Option Explicit

'---------------------------------------------------
'
'                     Lib_Misc
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Public Indent As New CIndent

'***************************************
'
'           Check file exists
'
'***************************************

Public Function FileExists(FileName As String) As Boolean
    Dim objFSO As New Scripting.FileSystemObject

    FileExists = objFSO.FileExists(FileName)

    Set objFSO = Nothing
End Function

'***************************************
'
'           Read a text file
'
'***************************************

Public Function ReadTextFile(Path As String) As String
    Dim FNum As Integer
    Dim Length As Long
    Dim Buffer As String

    On Error GoTo Done
    FNum = FreeFile
    Open Path For Binary Access Read As #FNum

    Length = LOF(FNum)
    Buffer = Space$(Length)     'make space for the file

    Get #FNum, , Buffer         'super fast file read

    Close #FNum
    ReadTextFile = Buffer
Done:
    On Error GoTo 0
End Function

'***************************************
'
'    Convert a dictionary to string
'
'***************************************

Public Function DictToString(DictIn As Scripting.Dictionary, OutputType As String) As String
    Dim OutputTxt As String, ValStr As String, Separ As String
    Dim Opt As Variant, ValD As Variant

    If DictIn Is Nothing Then
        DictToString = ""
        Exit Function
    End If

    If OutputType = "JSON" Then
        OutputTxt = "{"
        For Each Opt In DictIn.Keys()
            If OutputTxt <> "{" Then OutputTxt = OutputTxt & ","
            'If a string came in, put double quotes around it
            ValD = DictIn(Opt)
            Separ = ""
            If VarType(ValD) = vbString Then Separ = """"

            'Value: correct for comma decimal system if a value was supplied
            ValStr = ValD
            If VarType(ValD) <> vbString Then ValStr = Replace(ValStr, ",", ".")
            OutputTxt = OutputTxt & """" & Opt & """" & ":" & Separ & ValStr & Separ
        Next
        OutputTxt = OutputTxt & "}"
    ElseIf OutputType = "URLENC" Then
        OutputTxt = ""
        For Each Opt In DictIn.Keys()
            If OutputTxt <> "" Then OutputTxt = OutputTxt & "&"
            ValD = DictIn(Opt)
            ValStr = ValD
            If VarType(ValD) <> vbString Then ValStr = Replace(ValStr, ",", ".")
            OutputTxt = OutputTxt & Opt & "=" & ValStr
        Next
    Else
        OutputTxt = "UNKNOWN_TYPE"
    End If

    DictToString = OutputTxt
End Function

'***************************************
'
'         Sort a dictionary
'
'***************************************

Public Sub SortDictByKey(DictIn As Scripting.Dictionary, Optional SortAsc As Boolean = True)
    Dim ResDict As New Scripting.Dictionary
    Dim ArrayList As Object
    Dim Key As Variant, va As Variant

    If DictIn Is Nothing Then
        Exit Sub
    End If
    If DictIn.Count <= 1 Then
        Exit Sub
    End If

    Set ArrayList = CreateObject("System.Collections.ArrayList")

    For Each Key In DictIn.Keys
        ArrayList.Add Key
    Next Key
    ArrayList.Sort
    If SortAsc = False Then
        ArrayList.Reverse
    End If

    For Each va In ArrayList
        ResDict.Add va, DictIn(va)
    Next va

    Set DictIn = ResDict
End Sub
