Attribute VB_Name = "M_common"
Option Explicit
Option Private Module

Function dayOfWeek(jdn As Long, _
                 Optional style As Integer = ISO_8601) _
     As Integer
    dayOfWeek = ((jdn + style - 1) Mod 7) + 1
End Function


Function weekdayname(jdn As Long, _
                     Optional Language As Integer = 1) As String
    
    Dim Monday As String
    Dim Tuesday As String
    Dim Wednesday As String
    Dim Thursday As String
    Dim Friday As String
    Dim Saturday As String
    Dim Sunday As String
        
    Select Case Language
    Case Hebrew
         Monday = "yom sheni"
         Tuesday = "yom shlishi"
         Wednesday = "yom r'vi'i"
         Thursday = "yom xamishi"
         Friday = "yom shishi"
         Saturday = "yom shabat"
         Sunday = "yom rishon"
    Case Arabic
         Monday = "Al-`iTHnain"
         Tuesday = "Al-THulaTHa"
         Wednesday = "Al-Arbi'a"
         Thursday = "Al-KHamees"
         Friday = "Al-Jum'ah"
         Saturday = "As-Sabt"
         Sunday = "Al-Ahad"
    Case Dutch
         Monday = "Maandag"
         Tuesday = "Dinsdag"
         Wednesday = "Woensdag"
         Thursday = "Donderdag"
         ' Use y-umlaut in stead of "ij" so the short version
         ' of "Vrÿdag" will look like "Vrÿ" in stead of "Vri".
         ' The real Dutch "ij" character would be better, but
         ' the Chr-function refuse to return unicode character &h0133.
         Friday = "Vr" & ChrW(&HFF) & "dag"
         Saturday = "Zaterdag"
         Sunday = "Zondag"
    Case French
         Monday = "Lundi"
         Tuesday = "Mardi"
         Wednesday = "Mercredi"
         Thursday = "Jeudi"
         Friday = "Vendredi"
         Saturday = "Samedi"
         Sunday = "Dimanche"
    Case German
         Monday = "Montag"
         Tuesday = "Dienstag"
         Wednesday = "Mittwoch"
         Thursday = "Donnerstag"
         Friday = "Freitag"
         Saturday = "Samstag"
         Sunday = "Sonntag"
    Case Farsi
         Monday = "Doshanbeh"
         Tuesday = "Seshanbeh"
         Wednesday = "Chaharshanbeh"
         Thursday = "Panjshanbeh"
         Friday = "Jomeh"
         Saturday = "Shanbeh"   ' First day of the week in Persian Calendar
         Sunday = "Yekshanbeh"
    Case Else
         Monday = "Monday"
         Tuesday = "Tuesday"
         Wednesday = "Wednesday"
         Thursday = "Thursday"
         Friday = "Friday"
         Saturday = "Saturday"
         Sunday = "Sunday"
    End Select
    
    Select Case dayOfWeek(jdn, ISO_8601)
    Case 1
         weekdayname = Monday
    Case 2
         weekdayname = Tuesday
    Case 3
         weekdayname = Wednesday
    Case 4
         weekdayname = Thursday
    Case 5
         weekdayname = Friday
    Case 6
         weekdayname = Saturday
    Case Else
         weekdayname = Sunday
    End Select
End Function

