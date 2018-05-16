Attribute VB_Name = "M_persian"
Option Explicit
Option Private Module

'The modern Persian calendar was adopted in 1925, supplanting (while
'retaining the month names of) a traditional calendar dating from the
'eleventh century. The calendar  consists of 12 months, the first six
'of which are 31 days, the next five 30 days, and the  final month 29
'days in a normal year and 30 days in a leap year.
'
'As one of the few calendars designed in the era of accurate
'positional astronomy, the  Persian calendar uses a very complex leap
'year structure which makes it the most  accurate solar calendar in use
'today. Years are grouped into cycles which begin with four normal
'years after which every fourth subsequent year in the cycle is a leap
'year. Cycles are grouped into grand cycles of either 128 years
'(composed of cycles of 29, 33, 33, and 33 years) or 132 years,
'containing cycles of of 29, 33, 33, and 37 years. A great grand cycle
'is composed of 21 consecutive 128 year grand cycles and a final 132
'grand cycle, for a total of 2820 years. The pattern of normal and leap
'years which began in 1925 will not repeat until the year 4745!
'
'Each 2820 year great grand cycle contains 2137 normal years of 365
'days and 683 leap years of 366 days, with the average year length over
'the great grand cycle of 365.24219852. So close is this to the actual
'solar tropical year of 365.24219878 days that the Persian calendar
'accumulates an error of one day only every 3.8 million years. As a
'purely solar calendar, months are not synchronised with the phases of
'the Moon.
'
'The Persian calendar is also known as the Sun's Hijri calendar. In Iran
'the Persian calendar is called Shamsi.
'
'Thanks to Mohammad Tahani <mtahani@noornet.net> from the Computer
'Research Center of Islamic Sciences ( CRCIS ) of Qom
'(http://www.noornet.net) for his help.

' We needed an alternative to Int and Fix.
' Int(8.4) = 8, Int(-8.4) = -9
' Fix(8.4) = 8, Fix(-8.4) = -8
' Ceil(8.4) = 9, Ceil(-8.4) = -9
Private Function Ceil(number As Single) As Long
    Ceil = -Sgn(number) * Int(-Abs(number))
' or
    'Ceil = CInt(number + (Sgn(number) * 0.5))
End Function

' Is a given year a leap year in the Persian calendar ?
'
Private Function leap_persian(iYear As Integer) As Boolean
    If iYear > 0 Then
        leap_persian = ((((((iYear - (474)) Mod 2820) + 474) + 38) * 682) Mod 2816) < 682
    Else
        leap_persian = ((((((iYear - (473)) Mod 2820) + 474) + 38) * 682) Mod 2816) < 682
    End If
End Function
'
' Determine Julian day from Persian date
'
Function persian_jdn(iYear As Integer, _
                     iMonth As Integer, _
                     iDay As Integer) As Long
    Const PERSIAN_EPOCH = 1948321 ' The JDN of 1 Farvardin 1
    Dim epbase As Long
    Dim epyear As Long
    Dim mdays As Long
    If iYear >= 0 Then
        epbase = iYear - 474
    Else
        epbase = iYear - 473
    End If
    epyear = 474 + (epbase Mod 2820)
    If iMonth <= 7 Then
        mdays = (CLng(iMonth) - 1) * 31
    Else
        mdays = (CLng(iMonth) - 1) * 30 + 6
    End If
    persian_jdn = CLng(iDay) _
            + mdays _
            + Fix(((epyear * 682) - 110) / 2816) _
            + (epyear - 1) * 365 _
            + Fix(epbase / 2820) * 1029983 _
            + (PERSIAN_EPOCH - 1)
End Function
'
Sub jdn_persian(jdn As Long, _
                ByRef iYear As Integer, _
                ByRef iMonth As Integer, _
                ByRef iDay As Integer)
    Dim depoch
    Dim cycle
    Dim cyear
    Dim ycycle
    Dim aux1, aux2
    Dim yday
    depoch = jdn - persian_jdn(475, 1, 1)
    cycle = Fix(depoch / 1029983)
    cyear = depoch Mod 1029983
    If cyear = 1029982 Then
        ycycle = 2820
    Else
        aux1 = Fix(cyear / 366)
        aux2 = cyear Mod 366
        ycycle = Int(((2134 * aux1) + (2816 * aux2) + 2815) / 1028522) + aux1 + 1
    End If
    iYear = ycycle + (2820 * cycle) + 474
    If iYear <= 0 Then
        iYear = iYear - 1
    End If
    yday = (jdn - persian_jdn(iYear, 1, 1)) + 1
    If yday <= 186 Then
        iMonth = Ceil(yday / 31)
    Else
        iMonth = Ceil((yday - 6) / 30)
    End If
    iDay = (jdn - persian_jdn(iYear, iMonth, 1)) + 1
End Sub
Function persian_MonthName(iMonth As Integer) As String
    Select Case iMonth
    Case 1
        persian_MonthName = "Farvardin"
    Case 2
        persian_MonthName = "Ordibehesht"
    Case 3
        persian_MonthName = "Khordad"
    Case 4
        persian_MonthName = "Tir"
    Case 5
        persian_MonthName = "Mordad"
    Case 6
        persian_MonthName = "Shahrivar"
    Case 7
        persian_MonthName = "Mehr"
    Case 8
        persian_MonthName = "Aban"
    Case 9
        persian_MonthName = "Azar"
    Case 10
        persian_MonthName = "Dey"
    Case 11
        persian_MonthName = "Bahman"
    Case 12
        persian_MonthName = "Esfand"
    Case Else
        persian_MonthName = "Error"
    End Select
End Function
