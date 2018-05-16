VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«— »«ÿ »« „«"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12195
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   9135
      Left            =   0
      Picture         =   "Form16.frx":2CFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub
