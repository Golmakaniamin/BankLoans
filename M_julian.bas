Attribute VB_Name = "M_julian"
Option Explicit
Option Private Module
'
Sub jdn_julian(jdn As Long, _
                       ByRef iYear As Integer, _
                       ByRef iMonth As Integer, _
                       ByRef iDay As Integer)
    Dim l As Long
    Dim k As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long

    j = jdn + 1402
    k = ((j - 1) \ 1461)
    l = j - 1461 * k
    n = ((l - 1) \ 365) - (l \ 1461)
    i = l - 365 * n + 30
    j = ((80 * i) \ 2447)
    iDay = i - ((2447 * j) \ 80)
    i = (j \ 11)
    iMonth = j + 2 - 12 * i
    iYear = 4 * k + n + i - 4716

End Sub
'
'
'
'
Function julian_jdn(iYear As Integer, _
                    iMonth As Integer, _
                    iDay As Integer) As Long
    Dim lYear As Long
    Dim lMonth As Long
    Dim lDay As Long

    lYear = CLng(iYear)
    lMonth = CLng(iMonth)
    lDay = CLng(iDay)

    julian_jdn = 367 * lYear - _
            ((7 * (lYear + 5001 + ((lMonth - 9) \ 7))) \ 4) _
            + ((275 * lMonth) \ 9) + lDay + 1729777

End Function
'
'
'
'
Sub julian_normDate(ByRef iYear As Integer, _
                   ByRef iMonth As Integer, _
                   ByRef iDay As Integer)
    Call civil_normDate(iYear, iMonth, iDay, Julian)
End Sub
'
'
'
'
Function julian_monthName(iMonth As Integer) As String
    julian_monthName = civil_monthName(iMonth)
End Function
