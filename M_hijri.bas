Attribute VB_Name = "M_hijri"
Option Explicit
Option Private Module

' Module Hijri
'
' Some Code in this module is based on a C-program called hdate
' written by Waleed Muhanna <wmuhanna@magnus.acs.ohio-state.edu>
' Waleed Muhanna based his code partially on "Astronomical
' Formulae for Calculators" by Jean Meeus
'
' The routines in this module use one of the  simplest astronomical
' criterion for earlist lunar visibility.  The criterion is a
' local one and is based on the moon's age since the lunar
' conjunction (the start of the new moon).
'
' The minimum recorded age of the crescent moon for first
' visibility is around 13 1/2 hours.  These routines take this
' into consideration to determine the likelihood of the
' crescent being sighted just after the first sun set (assumed
' to occur at 7:30p.m.) following the new moon.  And since
' there is no agreed upon international Lunar date-line, hcal
' and hdate currently use Makkah (+3 GMT) for the locale.

'
' Given an islamic month number, return the name of the month
' For some months alternate Latin spellings have been provided
' as comment
'
Function islamic_MonthName(iMonth As Integer) As String
    Select Case iMonth
    Case 1
        islamic_MonthName = "Muharram"      ' or "MuHaram"
    Case 2
        islamic_MonthName = "Safar"
    Case 3
        islamic_MonthName = "Rabi' I"       ' or "Raby` al-awal"
    Case 4
        islamic_MonthName = "Rabi' II"      ' or "Raby` al-THaany"
    Case 5
        islamic_MonthName = "Jumada I"      ' or "Jumaada al-awal"
    Case 6
        islamic_MonthName = "Jumada II"     ' or "Jumaada al-THaany"
    Case 7
        islamic_MonthName = "Rajab"
    Case 8
        islamic_MonthName = "Sha'ban"       ' or "SHa`baan"
    Case 9
        islamic_MonthName = "Ramadan"       ' or "RamaDHaan"
    Case 10
        islamic_MonthName = "Shawwal"
    Case 11
        islamic_MonthName = "Dhu 'l-Qa'dah" ' or "Thw al-Qi`dah"
    Case 12
        islamic_MonthName = "Dhu 'l-Hijja"  ' or "Thw al-Hijjah"
    Case Else
        islamic_MonthName = "Error"
    End Select
End Function

'
' Given an integer _n_ and a phase selector (nph=0,1,2,3 for
' new,first,full,last quarters respectively, function returns the
' Julian date/time (integer part is the julian day number,
' fraction is the time) of the Nth such phase since January 1900.
' Adapted from "Astronomical  Formulae for Calculators" by
' Jean Meeus, Third Edition, Willmann-Bell, 1985.
'
Function tmoonphase(n As Long, _
                    nph As Integer) As Double
    
    Const RPD = (1.74532925199433E-02) ' radians per degree (pi/180)

    Dim jd As Double
    Dim t As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim k As Double
    Dim ma As Double
    Dim sa As Double
    Dim tf As Double
    Dim xtra As Double

    k = n + nph / 4#
    t = k / 1236.85
    t2 = t * t
    t3 = t2 * t
    jd = 2415020.75933 + 29.53058868 * k - 0.0001178 * t2 _
        - 0.000000155 * t3 _
        + 0.00033 * Sin(RPD * (166.56 + 132.87 * t - 0.009173 * t2))
'
'   Sun's mean anomaly
    sa = RPD * (359.2242 + 29.10535608 * k - 0.0000333 * t2 - 0.00000347 * t3)
'
'   Moon's mean anomaly
    ma = RPD * (306.0253 + 385.81691806 * k + 0.0107306 * t2 + 0.00001236 * t3)
    
'
'   Moon's argument of latitude
    tf = RPD * 2# * (21.2964 + 390.67050646 * k - 0.0016528 * t2 _
              - 0.00000239 * t3)
'
'   should reduce to interval 0-1.0 before calculating further
    Select Case nph
    Case 0, 2
        xtra = (0.1734 - 0.000393 * t) * Sin(sa) _
              + 0.0021 * Sin(sa * 2) _
              - 0.4068 * Sin(ma) + 0.0161 * Sin(2 * ma) - 0.0004 * Sin(3 * ma) _
              + 0.0104 * Sin(tf) _
              - 0.0051 * Sin(sa + ma) - 0.0074 * Sin(sa - ma) _
              + 0.0004 * Sin(tf + sa) - 0.0004 * Sin(tf - sa) _
              - 0.0006 * Sin(tf + ma) + 0.001 * Sin(tf - ma) _
              + 0.0005 * Sin(sa + 2 * ma)
    Case 1, 3
        xtra = (0.1721 - 0.0004 * t) * Sin(sa) _
              + 0.0021 * Sin(sa * 2) _
              - 0.628 * Sin(ma) + 0.0089 * Sin(2 * ma) - 0.0004 * Sin(3 * ma) _
              + 0.0079 * Sin(tf) _
              - 0.0119 * Sin(sa + ma) - 0.0047 * Sin(sa - ma) _
              + 0.0003 * Sin(tf + sa) - 0.0004 * Sin(tf - sa) _
              - 0.0006 * Sin(tf + ma) + 0.0021 * Sin(tf - ma) _
              + 0.0003 * Sin(sa + 2 * ma) + 0.0004 * Sin(sa - 2 * ma) _
              - 0.0003 * Sin(2 * sa + ma)
        If (nph = 1) Then
            xtra = xtra + 0.0028 - 0.0004 * Cos(sa) + 0.0003 * Cos(ma)
        Else
            xtra = xtra - 0.0028 + 0.0004 * Cos(sa) - 0.0003 * Cos(ma)
        End If
    Case Else
        tmoonphase = 0
        Exit Function
    End Select
'   convert from Ephemeris Time (ET) to (approximate)Universal Time (UT)
    tmoonphase = jd + xtra - (0.41 + 1.2053 * t + 0.4992 * t2) / 1440
End Function
'
' parameters for Makkah: for a new moon to be visible after sunset on
' a the same day in which it started, it has to have started before
' (SUNSET-MINAGE)-TIMZ=3 A.M. local time.
'
Function visibility(n As Long) As Double
    
    ' parameters for Makkah: for a new moon to be visible after sunset on
    ' a the same day in which it started, it has to have started before
    ' (SUNSET-MINAGE)-TIMZ=3 A.M. local time.
    Const TIMZ = 3#
    Const MINAGE = 13.5
    Const SUNSET = 19.5 ' approximate
    Const TIMDIF = (SUNSET - MINAGE)

    Dim jd As Double
    Dim tf As Single
    Dim d As Long
    
    jd = tmoonphase(n, 0)
    d = Int(jd)
    tf = (jd - d)
    If (tf <= 0.5) Then  ' new moon starts in the afternoon
        visibility = (jd + 1#)
    Else  ' new moon starts before noon
        tf = (tf - 0.5) * 24 + TIMZ ' local time
        If (tf > TIMDIF) Then
            visibility = (jd + 1#) ' age at sunset < min for visiblity
        Else
            visibility = (jd)
        End If
    End If
End Function
'
'
' Given a julian day number, compute corresponding Hijri date.
' As a reference point, the routine uses the fact that the iYear
' 1405 A.H. started immediatly after lunar conjunction number 1048
' which occured on September 1984 25d 3h 10m UT.
'
Sub jdn_islamic(jd As Long, _
              ByRef iYear As Integer, _
              ByRef iMonth As Integer, _
              ByRef iDay As Integer)
    Dim mjd As Double
    Dim k As Long
    Dim hm As Long

    Call jdn_civil(jd, iYear, iMonth, iDay)
    k = Int(0.6 + (iYear + (CInt(iMonth - 0.5)) / 12# + iDay / 365# - 1900) * 12.3685)
    Do
        mjd = visibility(k)
        k = k - 1
    Loop While (mjd > (jd - 0.5))
    k = k + 1
    hm = k - 1048
    iYear = 1405 + Fix(hm / 12)
    'iYear = 1405 + Int(hm / 12)
    
    iMonth = (hm Mod 12) + 1
    If (hm <> 0 And iMonth <= 0) Then
        iMonth = iMonth + 12
        iYear = iYear - 1
    End If
    If iYear <= 0 Then iYear = iYear - 1
    iDay = Int(jd - mjd + 0.5)
End Sub
'
' Given a Hijri date, compute corresponding julian day number
'
Function islamic_jdn(ByVal iYear As Integer, _
                   iMonth As Integer, _
                   iDay As Integer) As Long
    
    ' NMONTH is the number of months between julian day number 1 and
    ' the iYear 1405 A.H. which started immediatly after lunar
    ' conjunction number 1048 which occured on September 1984 25d
    ' 3h 10m UT.
    Const NMONTHS = (1405 * 12 + 1)

    Dim k As Long
    
    If (iYear < 0) Then iYear = iYear + 1
    k = iMonth + iYear * 12 - NMONTHS ' nunber of months since 1/1/1405
    islamic_jdn = Int(visibility(k + CLng(1048)) + iDay + 0.5)
End Function
