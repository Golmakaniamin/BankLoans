Attribute VB_Name = "M_constants"
Option Explicit
'
Public Const CalmathVersion = "Autumn 2001 build 920"
'
' default day- and weeknumstyle
Public Const ISO_8601 = 1
'
'Calendar types
Public Const Gregorian = ISO_8601
Public Const Julian = 2
Public Const Hebrew = 3
Public Const Islamic = 4
Public Const Persian = 98
'
' Weekday Numbering styles
Public Const weekStartsOnMonday = ISO_8601
Public Const weekStartsOnSunday = 2
Public Const weekStartsOnSaturday = 3
Public Const weekStartsOnFriday = 4
Public Const weekStartsOnThursday = 5
Public Const weekStartsOnWednesday = 6
Public Const weekStartsOnTuesday = 7
'
' weeknumstyles
Public Const Jan1InWeek1 = 0
Public Const Jan4InWeek1 = ISO_8601
Public Const Jan7InWeek1 = 2
'
' Languages
Public Const English = 1
Public Const Arabic = Islamic
'public const Hebrew = 3 ' allready defined as such above
Public Const Dutch = 31
Public Const French = 33
Public Const German = 49
Public Const Farsi = Persian
'
'Miscellaneous public constants
Public Const Signed = -1
Public Const UnSigned = 1
