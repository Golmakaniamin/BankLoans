Attribute VB_Name = "M_civil"
Option Explicit
Option Private Module
'
Sub jdn_civil(jdn As Long, _
              ByRef iYear As Integer, _
              ByRef iMonth As Integer, _
              ByRef iDay As Integer)

    Dim l As Long
    Dim k As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long

    If (jdn > 2299160) Then
        l = jdn + 68569
        n = ((4 * l) \ 146097)
        l = l - ((146097 * n + 3) \ 4)
        i = ((4000 * (l + 1)) \ 1461001)
        l = l - ((1461 * i) \ 4) + 31
        j = ((80 * l) \ 2447)
        iDay = l - ((2447 * j) \ 80)
        l = (j \ 11)
        iMonth = j + 2 - 12 * l
        iYear = 100 * (n - 49) + i + l
    Else
        Call jdn_julian(jdn, iYear, iMonth, iDay)
    End If

End Sub
'
'
'
'
'

Function civil_weekNumber(iYear As Integer, _
                    iMonth As Integer, _
                    iDay As Integer, _
                    Optional DaynumStyle As Integer = ISO_8601, _
                    Optional WeeknumStyle As Integer = ISO_8601, _
                    Optional CalendarStyle As Integer = Gregorian) As Integer
    Dim markday As Long
    Dim firstmarkday As Long

    Select Case WeeknumStyle
    Case Jan1InWeek1
        civil_weekNumber = (civil_jdn(iYear, iMonth, iDay, CalendarStyle) _
                   - civil_jdn(iYear, 1, 1, CalendarStyle) _
                   + dayOfWeek(civil_jdn(iYear, 1, 1, CalendarStyle), DaynumStyle) + 6) \ 7
    Case Jan4InWeek1, Jan7InWeek1
        markday = civil_jdn(iYear, iMonth, iDay, CalendarStyle) _
                   - dayOfWeek(civil_jdn(iYear, iMonth, iDay, CalendarStyle), DaynumStyle) _
                 + 1 + (WeeknumStyle * 3)
        If markday > civil_jdn(iYear, 12, 31, CalendarStyle) Then
            'this week's markday is next iYear, so we're already in week 1.
            civil_weekNumber = 1
        Else
            If markday < civil_jdn(iYear, 1, 1, CalendarStyle) Then
                'We're still is last iYear's last week.
                civil_weekNumber = civil_weekNumber(iYear - 1, 12, 31, _
                                        DaynumStyle, WeeknumStyle)
            Else
                'This is a normal week in the middle of the iYear.
                'Count the weeks between this week's markday and
                'the first markday of the iYear, divide it by seven
                'and add 1. That's it!
                firstmarkday = civil_jdn(iYear, 1, 1 + (WeeknumStyle * 3), _
                                         CalendarStyle) _
                             - dayOfWeek(civil_jdn(iYear, 1, 1 + (WeeknumStyle * 3), _
                                         CalendarStyle), DaynumStyle) _
                             + 1 + (WeeknumStyle * 3)
                civil_weekNumber = (markday - firstmarkday) \ 7 + 1
            End If
        End If
    Case Else
        civil_weekNumber = 0
    End Select
End Function
'
'
'
'

Sub civil_normDate(ByRef iYear As Integer, _
                   ByRef iMonth As Integer, _
                   ByRef iDay As Integer, _
                   Optional CalendarType As Integer = Gregorian)
    Dim monLength As Integer

    If (iMonth > 12) Then
        iYear = iYear + (iMonth \ 12)
        iMonth = iMonth Mod 12
    ElseIf (iMonth < 1) Then
        iYear = iYear + (iMonth \ 12) - 1
        iMonth = 12 + (iMonth Mod 12)
    End If

    monLength = civil_daysInMonth(iYear, iMonth, CalendarType)
    While (iDay > monLength)
        If (iMonth = 12) Then
            iMonth = 1
            iYear = iYear + 1
        Else
            iMonth = iMonth + 1
        End If
        iDay = iDay - monLength
        monLength = civil_daysInMonth(iYear, iMonth, CalendarType)
    Wend

    While (iDay < 1)
        If (iMonth = 1) Then
            iMonth = 12
            iYear = iYear - 1
        Else
            iMonth = iMonth - 1
        End If
        iDay = iDay + civil_daysInMonth(iYear, iMonth, CalendarType)
    Wend
End Sub
'
'
'
'
Function civil_monthName(iMonth As Integer, _
                         Optional Language As Integer = 1) As String
    Dim Jan As String
    Dim Feb As String
    Dim Mar As String
    Dim Apr As String
    Dim May As String
    Dim Jun As String
    Dim Jul As String
    Dim Aug As String
    Dim Sep As String
    Dim Oct As String
    Dim Nov As String
    Dim Dec As String
        
    Select Case Language
    Case Dutch
         Jan = "Januari"
         Feb = "Februari"
         Mar = "Maart"
         Apr = "April"
         May = "Mei"
         Jun = "Juni"
         Jul = "Juli"
         Aug = "Augustus"
         Sep = "September"
         Oct = "Oktober"
         Nov = "November"
         Dec = "December"
    Case French
         Jan = "Janvier"
         Feb = "Février"
         Mar = "Mars"
         Apr = "Avril"
         May = "Mai"
         Jun = "Juin"
         Jul = "Juillet"
         Aug = "Août"
         Sep = "Septembre"
         Oct = "Octobre"
         Nov = "Novembre"
         Dec = "Décembre"
    Case German
         Jan = "Januar"
         Feb = "Februar"
         Mar = "März"
         Apr = "April"
         May = "Mai"
         Jun = "Juni"
         Jul = "Juli"
         Aug = "August"
         Sep = "September"
         Oct = "Oktober"
         Nov = "November"
         Dec = "Dezember"
    Case Else
         Jan = "January"
         Feb = "February"
         Mar = "March"
         Apr = "April"
         May = "May"
         Jun = "June"
         Jul = "July"
         Aug = "August"
         Sep = "September"
         Oct = "October"
         Nov = "November"
         Dec = "December"
    End Select

    Select Case iMonth
    Case 1
        civil_monthName = Jan
    Case 2
        civil_monthName = Feb
    Case 3
        civil_monthName = Mar
    Case 4
        civil_monthName = Apr
    Case 5
        civil_monthName = May
    Case 6
        civil_monthName = Jun
    Case 7
        civil_monthName = Jul
    Case 8
        civil_monthName = Aug
    Case 9
        civil_monthName = Sep
    Case 10
        civil_monthName = Oct
    Case 11
        civil_monthName = Nov
    Case 12
        civil_monthName = Dec
    Case Else
        civil_monthName = "Error"
    End Select
End Function
'
'
'
'
Function civil_leapyear(iYear As Integer, _
                Optional CalendarType As Integer = Gregorian) _
         As Boolean

    If CalendarType <> Gregorian Or iYear < 1582 Then
    ' Julian calendar
        civil_leapyear = ((iYear Mod 4) = 0)
    Else
    ' Gregorian calendar
        civil_leapyear = ((iYear Mod 4) = 0) Xor _
                         ((iYear Mod 100) = 0) Xor _
                         ((iYear Mod 400) = 0)
    End If

End Function
'
'
'
'
Function civil_jdn(iYear As Integer, _
                   iMonth As Integer, _
                   iDay As Integer, _
                   Optional CalendarType As Integer = Gregorian) As Long
    Dim lYear As Long
    Dim lMonth As Long
    Dim lDay As Long

    If CalendarType = Gregorian And ((iYear > 1582) Or _
        ((iYear = 1582) And (iMonth > 10)) Or _
        ((iYear = 1582) And (iMonth = 10) And (iDay > 14))) _
    Then
        lYear = CLng(iYear)
        lMonth = CLng(iMonth)
        lDay = CLng(iDay)
        civil_jdn = ((1461 * (lYear + 4800 + ((lMonth - 14) \ 12))) \ 4) _
            + ((367 * (lMonth - 2 - 12 * (((lMonth - 14) \ 12)))) \ 12) _
            - ((3 * (((lYear + 4900 + ((lMonth - 14) \ 12)) \ 100))) \ 4) _
            + lDay - 32075
    Else
        civil_jdn = julian_jdn(iYear, iMonth, iDay)
    End If

End Function
'
'
'
'
Function civil_daysInMonth(iYear As Integer, _
                     iMonth As Integer, _
                     Optional CalendarType As Integer = Gregorian) As Integer
    Select Case iMonth
    Case 2 ' February
        If civil_leapyear(iYear, CalendarType) Then
            civil_daysInMonth = 29
        Else
            civil_daysInMonth = 28
        End If
    Case 4, 6, 9, 11 ' April, June, September or November
        civil_daysInMonth = 30
    Case Else    ' Other values.
        civil_daysInMonth = 31
    End Select
End Function

