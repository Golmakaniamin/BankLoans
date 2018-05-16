Attribute VB_Name = "Module1"
'---------------------------------------'

'   http://www.visualbasic.blogfa.com

'---------------------------------------'

Option Explicit

Private Const hezar = " Â“«— "
Private Const melun = " „Ì·ÌÊ‰ "
Private Const melyard = " „Ì·Ì«—œ "
Private Const va = " Ê "

'--- Farsi Number Convertor ------------------'

Public Function heji_adad(ByVal adad As Double) As String
    Dim hooroof As String
    Dim SS As Integer  'sadgan
    Dim hh As Integer   'hezargan
    Dim mm As Integer  'melungan
    Dim yy As Integer   'melyardgan
    Dim STRadad As String
    Dim LENadad As Integer
    
    STRadad = Str(Val(Str(adad)))
    LENadad = Len(STRadad)
    
    Select Case adad
                    Case Is = 0
                                    hooroof = "’›—"
                    Case 1 To 999
                                       hooroof = Adad_Heji(adad)
    
                    Case 1000 To 999999
    
                                If (adad Mod 1000 = 0) Then hooroof = Adad_Heji(Int(adad / 1000)) + hezar
                                If (adad Mod 1000 <> 0) Then hooroof = Adad_Heji(Int(adad / 1000)) + hezar + va + (Adad_Heji(adad Mod 1000))
    
                    Case 1000000 To 999999999
    
                    SS = Val(Right$(STRadad, 3))
                     hh = Val(Mid$(STRadad, LENadad - 5, 3))
                     mm = Val(Left$(STRadad, LENadad - 6))
    
                   If (SS = 0 And hh = 0) Then hooroof = Adad_Heji(mm) + melun
                   If (SS = 0 And hh <> 0) Then hooroof = Adad_Heji(mm) + melun + va + Adad_Heji(hh) + hezar
                   If (SS <> 0 And hh = 0) Then hooroof = Adad_Heji(mm) + melun + va + Adad_Heji(SS)
                   If (SS <> 0 And hh <> 0) Then hooroof = Adad_Heji(mm) + melun + va + Adad_Heji(hh) + hezar + va + Adad_Heji(SS)
    
                   Case 1000000000 To 999999999999#
    
                           SS = Val(Right$(STRadad, 3))
                           hh = Val(Mid$(STRadad, LENadad - 5, 3))
                           mm = Val(Mid$(STRadad, LENadad - 8, 3))
                           yy = Val(Left$(STRadad, LENadad - 9))
    
                                If (SS = 0 And hh = 0 And mm = 0) Then hooroof = Adad_Heji(yy) + melyard
                                If (SS = 0 And hh = 0 And mm <> 0) Then hooroof = Adad_Heji(yy) + melyard + va + Adad_Heji(mm) + melun
                                If (SS = 0 And hh <> 0 And mm <> 0) Then hooroof = Adad_Heji(yy) + melyard + va + Adad_Heji(mm) + melun + va + Adad_Heji(hh) + hezar
                                If (SS <> 0 And hh <> 0 And mm <> 0) Then hooroof = Adad_Heji(yy) + melyard + va + Adad_Heji(mm) + melun + va + Adad_Heji(hh) + hezar + va + Adad_Heji(SS)
    
                       Case Is > 999999999999#
                                   hooroof = "⁄œœ Ê«—œ ‘œÂ Œ«—Ã «“ „ÕœÊœÂ „Ì »«‘œ "
    
        End Select
        
       heji_adad = hooroof
End Function

Private Function Adad_Heji(ByVal adad As Integer) As String
    Dim yekan As Byte
    Dim dahgan As Byte
    Dim sadgan As Byte
    Dim behooroof As String
    
    Dim heji(19) As String
    Dim heji_dahgan(9) As String
    Dim heji_sadgan(9) As String
    '-------------------------------
    heji(1) = "Ìﬂ": heji(2) = "œÊ": heji(3) = "”Â": heji(4) = "çÂ«—": heji(5) = "Å‰Ã"
    heji(6) = "‘‘": heji(7) = "Â› ": heji(8) = "Â‘ ": heji(9) = "‰Â": heji(10) = "œÂ"
    heji(11) = "Ì«“œÂ": heji(12) = "œÊ«“œÂ": heji(13) = "”Ì“œÂ": heji(14) = "çÂ«—œÂ": heji(15) = "Å«‰“œÂ"
    heji(16) = "‘«‰“œÂ": heji(17) = "Â›œÂ": heji(18) = "ÂÌÃœÂ": heji(19) = "‰Ê“œÂ"
    '-------------------------------
    heji_dahgan(1) = "œÂ"
    heji_dahgan(2) = "»Ì”  "
    heji_dahgan(3) = "”Ì ": heji_dahgan(4) = "çÂ· ": heji_dahgan(5) = " Å‰Ã«Â"
    heji_dahgan(6) = "‘’  ": heji_dahgan(7) = "Â› «œ": heji_dahgan(8) = "Â‘ «œ"
    heji_dahgan(9) = "‰Êœ"
    '------------------------
    heji_sadgan(1) = "Ìﬂ’œ": heji_sadgan(2) = "œÊÌ” ": heji_sadgan(3) = "”Ì’œ"
    heji_sadgan(4) = "çÂ«—’œ": heji_sadgan(5) = "Å«‰’œ": heji_sadgan(6) = "‘‘’œ"
    heji_sadgan(7) = "Â› ’œ": heji_sadgan(8) = "Â‘ ’œ": heji_sadgan(9) = "‰Â’œ"
    '------------------------------------------------------------------------------------------------------------
             yekan = adad Mod 10
            dahgan = adad Mod 100
            sadgan = Int(adad / 100)
    '------------------------------------------------------------------------------------------------------------
            If dahgan < 20 Then
    
                        If (sadgan = 0) Then behooroof = heji(dahgan)
                        If (sadgan <> 0) Then behooroof = heji_sadgan(sadgan) + va + heji(dahgan)
                        If (yekan = 0 And dahgan = 0) Then behooroof = heji_sadgan(sadgan)
    
            Else
                dahgan = (adad Mod 100) - yekan
    
                           If (sadgan = 0 And yekan = 0) Then behooroof = heji_dahgan(dahgan / 10)
                           If (sadgan = 0 And yekan <> 0) Then behooroof = heji_dahgan(dahgan / 10) + va + heji(yekan)
                           If (sadgan <> 0 And yekan = 0) Then behooroof = heji_sadgan(sadgan) + va + heji_dahgan(dahgan / 10)
                           If (sadgan <> 0 And yekan <> 0) Then behooroof = heji_sadgan(sadgan) + va + heji_dahgan(dahgan / 10) + va + heji(yekan)
    
            End If
    
    Adad_Heji = behooroof
End Function

Private Function change_mony(ByVal Pol As Currency) As String
    Dim P As String, p1 As String
    Dim P2 As String, P3 As String
    Dim P4 As String, Sk As String
    Dim L As Byte
    If Pol > 0 Then
              P = Str(Pol)
              Sk = Right(Trim(P), 3)
              p1 = harf(Val(Trim(Sk)))
              p1 = Trim(p1) & " —Ì«·"
              If Len(Trim(P)) > 3 Then
                       Sk = Right(Trim(P), 6)
                       L = Len(Trim(Sk))
                       Sk = Left(Trim(Sk), (L - 3))
                       P2 = harf(Val(Trim(Sk)))
                       P2 = Trim(P2) & " Â“«— Ê"
              End If
              If Len(Trim(P)) > 6 Then
                       Sk = Right(Trim(P), 9)
                       L = Len(Trim(Sk))
                       Sk = Left(Trim(Sk), (L - 6))
                       P3 = harf(Val(Trim(Sk)))
                       P3 = Trim(P3) & " „Ì·ÌÊ‰ Ê"
              End If
              If Len(Trim(P)) = 10 Then
                       Sk = Left(Trim(P), 1)
                       P4 = harf(Val(Trim(Sk)))
                       P4 = Trim(P4) & " „Ì·Ì«—œ Ê"
              End If
              change_mony = Trim(P4) & Trim(P3) & Trim(P2) & Trim(p1)
    End If
End Function

Private Function harf(mony2 As Long) As String
    Dim S As String, S1 As String
    Dim s2 As String, s3 As String
    S = Trim(Str(mony2))
    If Len(Trim(Str(mony2))) = 1 Then S = "00" & Trim(Str(mony2))
    If Len(Trim(Str(mony2))) = 2 Then S = "0" & Trim(Str(mony2))
    Select Case Left(Trim(S), 1)
           Case 0
                 S1 = ""
           Case 1
                 S1 = "Ìﬂ’œ "
           Case 2
                 S1 = "œÊÌ”  "
           Case 3
                 S1 = "”Ì’œ"
           Case 4
                 S1 = "çÂ«—’œ"
           Case 5
                 S1 = "Å«‰’œ"
           Case 6
                 S1 = "‘‘’œ"
           Case 7
                 S1 = "Â› ’œ"
           Case 8
                 S1 = "Â‘ ’œ"
           Case 9
                 S1 = "‰Â’œ"
    End Select
    Select Case Mid(Trim(S), 2, 1)
           Case 0
                 s2 = ""
           Case 1
                 Select Case Right(Trim(S), 1)
                        Case 0
                             s2 = "œÂ"
                        Case 1
                             s2 = "Ì«“œÂ"
                        Case 2
                             s2 = "œÊ«“œÂ"
                        Case 3
                             s2 = "”Ì“œÂ"
                        Case 4
                             s2 = "çÂ«—œÂ"
                        Case 5
                             s2 = "Å«‰“œÂ"
                        Case 6
                             s2 = "‘«‰“œÂ"
                        Case 7
                             S1 = "Â›œÂ"
                        Case 8
                             S1 = "ÂÃœÂ"
                        Case 9
                             S1 = "‰Ê“œÂ"
                  End Select
           Case 2
                 s2 = "»Ì”  "
           Case 3
                 s2 = "”Ì "
           Case 4
                 s2 = " çÂ·"
           Case 5
                 s2 = " Å‰Ã«Â"
           Case 6
                 s2 = "‘’  "
           Case 7
                 s2 = "Â› «œ"
           Case 8
                 s2 = "Â‘ «œ"
           Case 9
                 s2 = "‰Êœ"
    End Select
    If Mid(Trim(S), 2, 1) <> 1 Then
        Select Case Right(Trim(S), 1)
               Case 0
                     s3 = ""
               Case 1
                     s3 = "Ìﬂ"
               Case 2
                     s3 = "œÊ"
               Case 3
                     s3 = "”Â"
               Case 4
                     s3 = "çÂ«—"
               Case 5
                     s3 = "Å‰Ã"
               Case 6
                     s3 = "‘‘"
               Case 7
                     s3 = "Â› "
               Case 8
                     s3 = "Â‘ "
               Case 9
                     s3 = "‰Â"
        End Select
    End If
    If Trim(S1) <> "" Then S1 = S1 & " Ê"
    If Trim(s2) <> "" Then s2 = s2 & " Ê"
    If Trim(s3) <> "" Then s3 = s3 & " Ê"
    S = S1 & s2 & s3
    If Trim(S) <> "" Then harf = Left(Trim(S), (Len(Trim(S)) - 1)) Else harf = ""
End Function
