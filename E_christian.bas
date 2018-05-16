Attribute VB_Name = "E_christian"
Option Explicit
Option Private Module
'
Sub pascha(iYear As Integer, _
           ByRef iMonth As Integer, _
           ByRef iDay As Integer, _
           Optional CalendarType As Integer = Gregorian)
    Dim c As Integer
    Dim g As Integer
    Dim k As Integer
    Dim i As Integer
    Dim j As Integer
    Dim l As Integer

    g = iYear Mod 19

    If CalendarType = Gregorian And iYear > 1582 Then
        c = iYear \ 100
        k = (c - 17) \ 25
        i = c - (c \ 4) - ((c - k) \ 3) + 19 * g + 15
        i = i Mod 30
        i = i - (i \ 28) * (1 - (i \ 28) * (29 \ (i + 1)) * ((21 - g) \ 11))
        j = iYear + (iYear \ 4) + i + 2 - c + (c \ 4)
        j = j Mod 7
    Else
        i = ((19 * g) + 15) Mod 30
        j = (iYear + (iYear \ 4) + i) Mod 7
    End If

    l = i - j
    iMonth = 3 + ((l + 40) \ 44)
    iDay = l + 28 - 31 * (iMonth \ 4)

End Sub
Function pascha_jdn(iYear As Integer, _
                    Optional CalendarType As Integer = Gregorian) As Long
    Dim iMonth As Integer
    Dim iDay As Integer
    
    Call pascha(iYear, iMonth, iDay, CalendarType)
    pascha_jdn = civil_jdn(iYear, iMonth, iDay, CalendarType)
End Function
'
'
'
'
Sub pentacost(iYear As Integer, _
              ByRef iMonth As Integer, _
              ByRef iDay As Integer, _
              Optional CalendarType As Integer = Gregorian)
    Call pascha(iYear, iMonth, iDay, CalendarType)
    iDay = iDay + 49
    civil_normDate iYear, iMonth, iDay, CalendarType
End Sub

