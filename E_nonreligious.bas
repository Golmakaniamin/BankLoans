Attribute VB_Name = "E_nonreligious"
Option Explicit
Option Private Module
'
Function nthWeekdayOfMonth(iYear As Integer, _
                            iMonth As Integer, _
                            WantedWeekday As Integer, _
                            n As Integer, _
                            Optional CalendarType As Integer = Gregorian) As Integer
Dim monthLength As Integer
    Select Case n
    Case 1 To 4
        nthWeekdayOfMonth = 1 + ((WantedWeekday - _
                                 dayOfWeek(civil_jdn(iYear, iMonth, 1, CalendarType)) + _
                                 7) Mod 7) + ((n - 1) * 7)
    Case -1, 5
        monthLength = civil_daysInMonth(iYear, iMonth, CalendarType)
        nthWeekdayOfMonth = monthLength - _
                            ((dayOfWeek(civil_jdn(iYear, iMonth, monthLength, CalendarType)) - _
                              WantedWeekday + 7) Mod 7)
    Case Else
        nthWeekdayOfMonth = 0
    End Select
End Function
