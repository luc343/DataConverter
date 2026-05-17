Attribute VB_Name = "Lib_JSON"
Option Explicit

'---------------------------------------------------
'
'                     Lib_Json
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

Public JsonConverter As New JsonConverter   'json parser

Public JsonConverterErr As Boolean

Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, ByRef lpLocalTime As SYSTEMTIME, ByRef lpUniversalTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

'***************************************
'
'      Convert local date to UTC
'
'***************************************

Public Function LocalDateToUTC(DT As Date) As String
    LocalDateToUTC = DateToUTCStr(LocalToUTC(DT))
End Function

Private Function LocalToUTC(ByVal tDate As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    Dim lRes As Long

    lRes = GetTimeZoneInformation(tzi)

    stLocal.wYear = Year(tDate)
    stLocal.wMonth = Month(tDate)
    stLocal.wDay = Day(tDate)
    stLocal.wHour = Hour(tDate)
    stLocal.wMinute = Minute(tDate)
    stLocal.wSecond = Second(tDate)
    stLocal.wMilliseconds = 0
    lRes = TzSpecificLocalTimeToSystemTime(tzi, stLocal, stUTC)

    LocalToUTC = DateSerial(stUTC.wYear, stUTC.wMonth, stUTC.wDay) + TimeSerial(stUTC.wHour, stUTC.wMinute, stUTC.wSecond)
End Function

Private Function DateToUTCStr(DT As Date) As String
    DateToUTCStr = Format$(DT, "yyyy-mm-ddTHH:nn:ss.000Z")
End Function
