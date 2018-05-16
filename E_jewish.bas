Attribute VB_Name = "E_jewish"
Option Explicit
Option Private Module
'
Function HebrewDateInOrAfterCivilYear(civilYear As Integer, _
              hebrewMonth As Integer, _
              hebrewDay As Integer, _
              Optional CalendarType As Integer = Gregorian) As Long
    Dim jdnJanuary1 As Long
    Dim jdnHoliday As Long
    Dim hebrewYear As Integer
    Dim dummy1 As Integer
    Dim dummy2 As Integer
    jdnJanuary1 = civil_jdn(civilYear, 1, 1, CalendarType)
    Call jdn_hebrew(jdnJanuary1, hebrewYear, dummy1, dummy2)
    jdnHoliday = hebrew_jdn(hebrewYear, hebrewMonth, hebrewDay)
    If jdnHoliday < jdnJanuary1 Then
        ' Oops! Wrong civil Year. Use next hebrewYear's in stead.
        jdnHoliday = hebrew_jdn(hebrewYear + 1, hebrewMonth, hebrewDay)
    End If
    HebrewDateInOrAfterCivilYear = jdnHoliday
End Function
'
'
'
'
Function HebrewDateInOrBeforeCivilYear(civilYear As Integer, _
              hebrewMonth As Integer, _
              hebrewDay As Integer, _
              Optional CalendarType As Integer = Gregorian) As Long
    Dim jdnDecember31 As Long
    Dim jdnHoliday As Long
    Dim hebrewYear As Integer
    Dim dummy1 As Integer
    Dim dummy2 As Integer
    jdnDecember31 = civil_jdn(civilYear, 12, 31, CalendarType)
    Call jdn_hebrew(jdnDecember31, hebrewYear, dummy1, dummy2)
    jdnHoliday = hebrew_jdn(hebrewYear, hebrewMonth, hebrewDay)
    If jdnHoliday > jdnDecember31 Then
        ' Oops! Wrong civil Year. Use last hebrewYear's in stead.
        jdnHoliday = hebrew_jdn(hebrewYear - 1, hebrewMonth, hebrewDay)
    End If
    HebrewDateInOrBeforeCivilYear = jdnHoliday
End Function

