Attribute VB_Name = "E_islamic"
Option Explicit
Option Private Module
'
Function FirstIslamicDateInCivilYear(civilYear As Integer, _
              islamicMonth As Integer, _
              islamicDay As Integer, _
              Optional CalendarType As Integer = Gregorian) As Long
    Dim jdnJanuary1 As Long
    Dim jdnHoliday As Long
    Dim islamicYear As Integer
    Dim dummy1 As Integer
    Dim dummy2 As Integer
    jdnJanuary1 = civil_jdn(civilYear, 1, 1, CalendarType)
    Call jdn_islamic(jdnJanuary1, islamicYear, dummy1, dummy2)
    jdnHoliday = islamic_jdn(islamicYear, islamicMonth, islamicDay)
    If jdnHoliday < jdnJanuary1 Then
        ' Oops! Wrong civil year. Use next islamicYear's in stead.
        jdnHoliday = islamic_jdn(islamicYear + 1, islamicMonth, islamicDay)
    End If
    FirstIslamicDateInCivilYear = jdnHoliday
End Function
'
'
'
'
Function LastIslamicDateInCivilYear(civilYear As Integer, _
             islamicMonth As Integer, _
             islamicDay As Integer, _
             Optional CalendarType As Integer = Gregorian) As Long
    Dim jdnDecember31 As Long
    Dim jdnHoliday As Long
    Dim islamicYear As Integer
    Dim dummy1 As Integer
    Dim dummy2 As Integer
    jdnDecember31 = civil_jdn(civilYear, 12, 31, CalendarType)
    Call jdn_islamic(jdnDecember31, islamicYear, dummy1, dummy2)
    jdnHoliday = islamic_jdn(islamicYear, islamicMonth, islamicDay)
    If jdnHoliday > jdnDecember31 Then
        ' Oops! Wrong civil year. Use last islamicYear's in stead.
        jdnHoliday = islamic_jdn(islamicYear - 1, islamicMonth, islamicDay)
    End If
    LastIslamicDateInCivilYear = jdnHoliday
End Function
