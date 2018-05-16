Attribute VB_Name = "E_Persian"
Option Explicit
Option Private Module
'
' In the Persian or Shamsi calendar, the following dates are holidays.
' * The first four days of each year ( Novrooz )
' * The "13th Farvardin" ( Sizdah )
' * All Fridays
' Furthermore all Islamic events are holidays in Iran.
' To determine the Shamsi date of an Islamic event, first
' determine the Islamic date of the event and then convert it
' to a Shamsi date.
' Since the Shamsi year is approximately 11 days longer than the
' Islamic year, Islamic events can occur twice within a Shamsi year.
' E.g. Qorban is celebrated on the Islamic date 10 Dhu 'l-Hijja.
' In the Shamsi year 1380 this corresponds to the 4th of Esfand,
' which is equivalent to the Civil date Februari 23, 2002.
' But in the Shamsi year 1411 Qorban will be celebrated on the
' 3rd of Farvardin (March 22, 2032) AND on the 22nd of Esfand
'(March 12, 2033)
'
Function FirstIslamicDateInPersianYear(persianYear As Integer, _
              islamicMonth As Integer, _
              islamicDay As Integer) As Long
    Dim jdnFarvardin1 As Long
    Dim jdnHoliday As Long
    Dim islamicYear As Integer
    Dim dummy1 As Integer
    Dim dummy2 As Integer
    jdnFarvardin1 = persian_jdn(persianYear, 1, 1)
    Call jdn_islamic(jdnFarvardin1, islamicYear, dummy1, dummy2)
    jdnHoliday = islamic_jdn(islamicYear, islamicMonth, islamicDay)
    If jdnHoliday < jdnFarvardin1 Then
        ' Oops! Wrong persian year. Use next islamicYear's in stead.
        jdnHoliday = islamic_jdn(islamicYear + 1, islamicMonth, islamicDay)
    End If
    FirstIslamicDateInPersianYear = jdnHoliday
End Function
'
Function LastIslamicDateInPersianYear(persianYear As Integer, _
             islamicMonth As Integer, _
             islamicDay As Integer) As Long
    Dim jdnEsfandLast As Long
    Dim jdnHoliday As Long
    Dim islamicYear As Integer
    Dim dummy1 As Integer
    Dim dummy2 As Integer
    ' The last day of Esfand can be Esfand 29 or Esfand 30, depending
    ' on whether it's a leap year or not. To determine the jdn of the
    ' last day of Esfand, determine the JDN of Farvardin 1 of the
    ' following year and subtract 1.
    jdnEsfandLast = persian_jdn(persianYear + 1, 1, 1) - 1
    Call jdn_islamic(jdnEsfandLast, islamicYear, dummy1, dummy2)
    jdnHoliday = islamic_jdn(islamicYear, islamicMonth, islamicDay)
    If jdnHoliday > jdnEsfandLast Then
        ' Oops! Wrong persian year. Use last islamicYear's in stead.
        jdnHoliday = islamic_jdn(islamicYear - 1, islamicMonth, islamicDay)
    End If
    LastIslamicDateInPersianYear = jdnHoliday
End Function

