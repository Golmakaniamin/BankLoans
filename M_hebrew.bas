Attribute VB_Name = "M_hebrew"
Option Explicit
Option Private Module
'
Sub jdn_hebrew(jdn As Long, _
               ByRef iYear As Integer, _
               ByRef iMonth As Integer, _
               ByRef iDay As Integer, _
               Optional monthcoding As Integer = UnSigned)
    Dim InputJDN As Long
    Dim tishri1 As Long
    Dim LeftOverDays As Long
    If jdn <= 347997 Then
        iYear = 0
        iMonth = 0
        iDay = 0
    Else
        InputJDN = jdn - 347997
        iYear = (InputJDN \ 365) + 1
        tishri1 = Hebrew_ElapsedCalendarDays(iYear)
        While (tishri1 > InputJDN)
            iYear = iYear - 1
            tishri1 = Hebrew_ElapsedCalendarDays(iYear)
        Wend
        iMonth = 1
        LeftOverDays = InputJDN - tishri1
        While (LeftOverDays >= Hebrew_LastDayOfMonth(iYear, iMonth))
            LeftOverDays = LeftOverDays - Hebrew_LastDayOfMonth(iYear, iMonth)
            iMonth = iMonth + 1
        Wend
        If Sgn(monthcoding) = Signed Then
            If iMonth > 6 Then
                If Hebrew_LeapYear(iYear) Then
                    iMonth = iMonth - 14
                Else
                    iMonth = iMonth - 13
                End If
            End If
        End If
        iDay = LeftOverDays + 1
    End If
End Sub
'
'
'
'
Function Hebrew_ShortKislev(iYear)
    Hebrew_ShortKislev = ((Hebrew_DaysInYear(iYear) Mod 10) = 3)
End Function
'
'
'
'
Function Hebrew_MonthName(iYear, ByVal iMonth) As String
        
    If ((iMonth > 6) And (Not (Hebrew_LeapYear(iYear)))) Then
       iMonth = iMonth + 1
    End If
    Select Case iMonth
    Case 1
        Hebrew_MonthName = "Tishri"     ' or "Tishrey"
    Case 2
        Hebrew_MonthName = "Heshvan"
    Case 3
        Hebrew_MonthName = "Kislev"
    Case 4
        Hebrew_MonthName = "Teveth"     ' or "Tevet"
    Case 5
        Hebrew_MonthName = "Shevat"
    Case 6
        If (Not (Hebrew_LeapYear(iYear))) Then
            Hebrew_MonthName = "Adar"
        Else
            Hebrew_MonthName = "Adar I"
        End If
    Case 7
        Hebrew_MonthName = "Adar II"
    Case 8
        Hebrew_MonthName = "Nisan"
    Case 9
        Hebrew_MonthName = "Iyyar"
    Case 10
        Hebrew_MonthName = "Sivan"
    Case 11
        Hebrew_MonthName = "Tammuz"
    Case 12
        Hebrew_MonthName = "Av"
    Case 13
        Hebrew_MonthName = "Elul"
    Case Else
        Hebrew_MonthName = ""
    End Select
End Function
'
'
'
'
Function Hebrew_LongHeshvan(iYear) As Boolean
    Hebrew_LongHeshvan = ((Hebrew_DaysInYear(iYear) Mod 10) = 5)
End Function
'
'
'
'
Function Hebrew_LeapYear(iYear) As Boolean
    If ((((7 * iYear) + 1) Mod 19) < 7) Then
       Hebrew_LeapYear = True
    Else
       Hebrew_LeapYear = False
    End If
End Function
'
'
'
'
Function Hebrew_LastDayOfMonth(iYear, ByVal iMonth) As Integer
    If ((iMonth > 6) And (Not (Hebrew_LeapYear(iYear)))) Then
       iMonth = iMonth + 1
    End If
    Select Case iMonth
    Case 2
        If Hebrew_LongHeshvan(iYear) Then
            Hebrew_LastDayOfMonth = 30
        Else
            Hebrew_LastDayOfMonth = 29
        End If
    Case 3
        If Hebrew_ShortKislev(iYear) Then
            Hebrew_LastDayOfMonth = 29
        Else
            Hebrew_LastDayOfMonth = 30
        End If
    Case 6
        If Hebrew_LeapYear(iYear) Then
            Hebrew_LastDayOfMonth = 30
        Else
            Hebrew_LastDayOfMonth = 29
        End If
    Case 4, 7, 9, 11, 13
        Hebrew_LastDayOfMonth = 29
    Case Else
        Hebrew_LastDayOfMonth = 30
    End Select
End Function
'
'
'
'
Function hebrew_jdn(iYear, ByVal iMonth, iDay) As Long
    Dim jdn As Long
    Dim counter As Integer
    If iMonth < 0 Then
        If Hebrew_LeapYear(iYear) Then
            iMonth = 14 + iMonth
        Else
            iMonth = 13 + iMonth
        End If
    End If
    jdn = Hebrew_ElapsedCalendarDays(iYear)
    For counter = 1 To (iMonth - 1) Step 1
        jdn = jdn + Hebrew_LastDayOfMonth(iYear, counter)
    Next counter
    hebrew_jdn = jdn + (iDay - 1 + 347997)
End Function
'
'
'
'
Function Hebrew_ElapsedCalendarDays(iYear) As Long
    Dim MonthsElapsed As Long
    Dim PartsElapsed As Long
    Dim HoursElapsed As Long
    Dim ConjunctionDay As Long
    Dim ConjunctionParts As Long
    Dim AlternativeDay As Long

    MonthsElapsed = (235 * (((iYear - 1) \ 19))) + _
                    (12 * ((iYear - 1) Mod 19)) + _
                    (7 * ((iYear - 1) Mod 19) + 1) \ 19
    PartsElapsed = 204 + 793 * (MonthsElapsed Mod 1080)
    HoursElapsed = 5 + 12 * MonthsElapsed + _
                   793 * ((MonthsElapsed \ 1080)) + _
                   PartsElapsed \ 1080
    ConjunctionDay = 1 + 29 * MonthsElapsed + HoursElapsed \ 24
    ConjunctionParts = (1080 * (HoursElapsed Mod 24)) + _
                       PartsElapsed Mod 1080
    If ((ConjunctionParts >= 19440) Or _
        (((ConjunctionDay Mod 7) = 2) And _
        (ConjunctionParts >= 9924) And _
        (Not (Hebrew_LeapYear(iYear)))) Or _
       (((ConjunctionDay Mod 7) = 1) And _
        (ConjunctionParts >= 16789) And _
        (Hebrew_LeapYear(iYear - 1)))) _
    Then
        AlternativeDay = ConjunctionDay + 1
    Else
        AlternativeDay = ConjunctionDay
    End If
    If (((AlternativeDay Mod 7) = 0) Or _
        ((AlternativeDay Mod 7) = 3) Or _
        ((AlternativeDay Mod 7) = 5)) _
    Then
        AlternativeDay = AlternativeDay + 1
    End If
    Hebrew_ElapsedCalendarDays = AlternativeDay
End Function
'
'
'
'
Function Hebrew_DaysInYear(iYear)
    Hebrew_DaysInYear = Hebrew_ElapsedCalendarDays(iYear + 1) - _
                        Hebrew_ElapsedCalendarDays(iYear)
End Function

