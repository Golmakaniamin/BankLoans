Attribute VB_Name = "M_convert"
Option Explicit
Option Private Module
'
' From civil to ...
'
Sub civil_julian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_julian(civil_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub civil_hebrew(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_hebrew(civil_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub civil_islamic(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_islamic(civil_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub civil_persian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_persian(civil_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
'
' From julian to ...
'
Sub julian_civil(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_civil(julian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub julian_hebrew(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_hebrew(julian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub julian_islamic(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_islamic(julian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub julian_persian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_persian(julian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
'
' From hebrew to ...
'
Sub hebrew_civil(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_civil(hebrew_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub hebrew_julian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_julian(hebrew_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub hebrew_islamic(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_islamic(hebrew_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub hebrew_persian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_persian(hebrew_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
'
' From islamic to ...
'
Sub islamic_civil(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_civil(islamic_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub islamic_julian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_julian(islamic_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub islamic_hebrew(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_hebrew(islamic_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub islamic_persian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_persian(islamic_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
'
' From persian to ...
'
Sub persian_civil(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_civil(persian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub persian_julian(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_julian(persian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub persian_hebrew(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_hebrew(persian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
Sub persian_islamic(ByRef iYear As Integer, _
                 ByRef iMonth As Integer, _
                 ByRef iDay As Integer)
    Call jdn_islamic(persian_jdn(iYear, _
                              iMonth, _
                              iDay), _
                    iYear, iMonth, iDay)
End Sub
