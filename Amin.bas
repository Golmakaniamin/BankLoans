Attribute VB_Name = "Amin"
Dim strd As String, d1 As String, m1 As String, y1 As String, d2 As String, m2 As String, y2 As String, strdate1 As String
Dim yt As String, mt As String, dt As String, sum As String

Public Function dateaminEktelaf(date1 As String, date2 As String) As String
If date1 <= date2 Then
  y1 = Mid(date1, 1, 4)
  m1 = Mid(date1, 6, 2)
  d1 = Mid(date1, 9, 2)

  y2 = Mid(date2, 1, 4)
  m2 = Mid(date2, 6, 2)
  d2 = Mid(date2, 9, 2)

  yt = y2 - y1
  If yt = 0 Then
    mt = m2 - m1
    If mt = 0 Then
      dateaminEktelaf = d2 - d1
    Else
      If mt < 6 Then dateaminEktelaf = ((mt - 1) * 30) + (d2 + (30 - d1))
      If mt > 6 Then dateaminEktelaf = ((mt - 1) * 30) + (d2 + (30 - d1))
    End If
  Else
    dateaminEktelaf = ((yt - 1) * 365) + (((m2 + (12 - m1)) - 1) * 30) + (d2 + (30 - d1))
  End If
Else
dateaminEktelaf = "a"
End If
End Function

Public Function dateaminEktelafmoon(date1 As String, date2 As String) As String
If date1 <= date2 Then
  y1 = Mid(date1, 1, 4)
  m1 = Mid(date1, 6, 2)

  y2 = Mid(date2, 1, 4)
  m2 = Mid(date2, 6, 2)

  yt = y2 - y1
  If yt = 0 Then
    mt = m2 - m1
    If mt <> 0 Then dateaminEktelafmoon = mt
  Else
    dateaminEktelafmoon = ((yt - 1) * 12) + ((m2 + (12 - m1)))
  End If
Else
  dateaminEktelafmoon = "a"
End If
End Function

Public Function dateaminEzafeMoon(date1 As String, number As String) As String
y1 = Mid(date1, 1, 4)
m1 = Mid(date1, 6, 2)
d1 = Mid(date1, 9, 2)

m1 = Val(m1) + number
If m1 > 12 Then
  y1 = Val(y1) + Val(m1 \ 12)
  m1 = m1 Mod 12
  If m1 = 0 Then
    y1 = y1 - 1
    m1 = 12
  End If
End If
If Len(m1) = 1 Then m1 = "0" + m1
dateaminEzafeMoon = y1 + "/" + m1 + "/" + d1
End Function

Public Function dateaminEzafeday(date1 As String, number As String) As String
y1 = Mid(date1, 1, 4)
m1 = Mid(date1, 6, 2)
d1 = Mid(date1, 9, 2)

d1 = Val(d1) + number

If d1 > 30 Then
  m1 = Val(m1) + Val(d1 \ 30)
  d1 = d1 Mod 30
End If

If m1 > 12 Then
  y1 = Val(y1) + Val(m1 \ 12)
  m1 = m1 Mod 12
  If m1 = 0 Then
    y1 = y1 - 1
    m1 = 12
  End If
End If

If Len(d1) = 1 Then d1 = "0" + d1
If Len(m1) = 1 Then m1 = "0" + m1
dateaminEzafeday = y1 + "/" + m1 + "/" + d1
End Function

Public Function moneyaminjoda(number As String) As String
Dim q As String
Dim w As Integer
q = ""
For w = 1 To Len(number)
  If Mid(number, w, 1) <> "." Then
    q = q + Mid(number, w, 1)
  End If
Next w
number = q
q = ""
e = Len(number) Mod 3

For w = Len(number) + 1 To 1 Step -3
  q = Mid(number, w, 3) + "." + q
Next w

q = Left(number, e) + "." + q
If Left(q, 1) = "." Then q = Right(q, Len(q) - 1)
If Right(q, 2) = ".." Then q = Left(q, Len(q) - 2)
moneyaminjoda = q
End Function

Public Function moneyaminnojoda(number As String) As Long
Dim q As String
q = ""
For w = 1 To Len(number)
  If Mid(number, w, 1) <> "." Then
    q = q + Mid(number, w, 1)
  End If
Next w
moneyaminnojoda = q
End Function

Function mil2shams(miladi_mm_dd_yyyy As String) As String
Dim iran(12), CHRIS(12)
CHRIS(1) = 31: CHRIS(2) = 28: CHRIS(3) = 31
CHRIS(4) = 30: CHRIS(5) = 31: CHRIS(6) = 30
CHRIS(7) = 31: CHRIS(8) = 31: CHRIS(9) = 30
CHRIS(10) = 31: CHRIS(11) = 30: CHRIS(12) = 31
For i = 1 To 12: iran(i) = 31 - (i \ 7) - (i \ 12): Next
mo = Val(Left(miladi_mm_dd_yyyy, 2))
miladi_mm_dd_yyyyy1 = Val(Mid(miladi_mm_dd_yyyy, 4, 2))
Year1 = Val(Mid(miladi_mm_dd_yyyy, 7, 4))
leap1 = Int((Year1 - 1) / 400)
leap2 = Year1 - 1 - 400 * leap1
leap3 = leap2 \ 100
leap4 = leap2 Mod 100
leap5 = leap4 \ 4
CHRIS(2) = 28
If ((Year1 Mod 4) = 0 And (Year1 Mod 100) <> 0) Or _
(Year1 Mod 400) = 0 Then CHRIS(2) = 29
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy1
For i = 1 To mo - 1
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy11 + CHRIS(i)
Next i
miladi_mm_dd_yyyyy1num = 365 * (Year1 - 1) + _
miladi_mm_dd_yyyyy11 + 97 * leap1 + 24 * leap3 + leap5
miladi_mm_dd_yyyyy1num = miladi_mm_dd_yyyyy1num - 221056!
iry1 = Int(miladi_mm_dd_yyyyy1num / 12053)
iry2 = miladi_mm_dd_yyyyy1num - 12053 * iry1
iry = 33 * iry1 - 16
If iry2 > 365 Then iry = iry + 1: iry2 = iry2 - 365
iry3 = iry2 \ 1461
iry4 = iry2 Mod 1461
iry5 = iry4 \ 365
iry6 = iry4 Mod 365
iry = iry + 1 + 4 * iry3 + iry5
iran(12) = 29
esfand = (8 * iry + 22) / 33 - 0.001
esfand = esfand - Int(esfand)
If esfand > 0.77 Then iran(12) = 30
For i = 1 To 12
If iry6 > iran(i) Then iry6 = iry6 - iran(i) _
Else irm = i: miladi_mm_dd_yyyyy11 = iry6: Exit For
Next i
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy11 + 5
If miladi_mm_dd_yyyyy11 > iran(irm) Then
miladi_mm_dd_yyyyy11 = miladi_mm_dd_yyyyy11 - iran(irm)
irm = irm + 1
If irm > 12 Then irm = 1: iry = iry + 1
End If
eirmiladi_mm_dd_yyyye = 3 * irm - 3
If irm > 7 Then eirmiladi_mm_dd_yyyye = _
eirmiladi_mm_dd_yyyye - irm + 7
girmiladi_mm_dd_yyyye = (8 * iry + 22) / 33 - 0.001
cirmiladi_mm_dd_yyyye = Int(girmiladi_mm_dd_yyyye) _
+ iry + eirmiladi_mm_dd_yyyye - miladi_mm_dd_yyyyy11 + 3
cirmiladi_mm_dd_yyyye = cirmiladi_mm_dd_yyyye Mod 7
If irm < 10 Then mo = "0" + LTrim(Str(irm)) Else _
mo = LTrim(Str(irm))
If miladi_mm_dd_yyyyy11 < 10 Then d = "0" + _
LTrim(Str(miladi_mm_dd_yyyyy11)) Else _
d = LTrim(Str(miladi_mm_dd_yyyyy11))
mil2shams = LTrim(Str(iry)) + "/" + mo + "/" + d
End Function

Public Function changeformatedate(strdate1 As String, amal As String, strdate2 As String) As String
  S1 = Mid(strdate, 1, 2)
  s2 = Mid(strdate, 4, 2)
  s3 = Mid(strdate, 7, 10)
  S = s2 + "-" + S1 + "-" + s3
  S = CDate(S) - CDate("05-02-2009")
  S1 = Mid(S, 1, 2)
  s2 = Mid(S, 4, 2)
  s3 = Mid(S, 7, 10)
  changeformatedate = s2 + "-" + S1 + "-" + s3
End Function

Public Function convertnumber(strdate As String) As String
  S1 = Mid(strdate, 1, 4)
  s2 = Mid(strdate, 6, 2)
  s3 = Mid(strdate, 9, 2)
  If s3 = "" Then
    convertnumber = S1 + "0" + s2
  Else
    convertnumber = S1 + "0" + s2 + "0" + s3
  End If
End Function

Public Function shams2mil(date1 As String) As String
Dim iShamsiYear As Integer
Dim iShamsiMonth As Integer
Dim iShamsiDay As Integer
Dim aa1 As String, aa2 As String

iShamsiYear = CInt(Mid(date1, 1, 4))
iShamsiMonth = CInt(Mid(date1, 6, 2))
iShamsiDay = CInt(Mid(date1, 9, 2))

persian_civil iShamsiYear, iShamsiMonth, iShamsiDay

If Len(CStr(iShamsiMonth)) = 1 Then
  aa1 = "0" + CStr(iShamsiMonth)
Else
  aa1 = CStr(iShamsiMonth)
End If

If Len(CStr(iShamsiDay)) = 1 Then
  aa2 = "0" + CStr(iShamsiDay)
Else
  aa2 = CStr(iShamsiDay)
End If

shams2mil = CStr(iShamsiYear) + "/" + aa1 + "/" + aa2
End Function
