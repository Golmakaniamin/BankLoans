VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFE4CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‰ŸÌ„« "
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE4CC&
      Caption         =   " ÂÌÂ ‰”ŒÂ Å‘ Ì»«‰"
      Height          =   4455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   3120
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "–ŒÌ—Â"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   3720
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Menu setting 
      Caption         =   " ‰ŸÌ„« "
      Begin VB.Menu mnubackup 
         Caption         =   "‰”ŒÂ Å‘ Ì»«‰"
         Begin VB.Menu mnusave 
            Caption         =   "–ŒÌ—Â"
         End
      End
      Begin VB.Menu mnugroup 
         Caption         =   " ⁄—Ì› ê—ÊÂ"
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub Command1_Click()

If fso.FolderExists(Dir1.Path + "\" + Text1.Text) = False Then
  fso.CreateFolder Dir1.Path + "\" + Text1.Text
  fso.CopyFile App.Path + "\1.mdb", Dir1.Path + "\" + Text1.Text + "\1.mdb"
Else
  MsgBox "„”Ì— Ê«—œ ‘œÂ ÊÃÊœ œ«—œ", vbCritical + vbMsgBoxRight
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
Frame1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub

Private Sub mnugroup_Click()
Form17.Show
End Sub

Private Sub mnusave_Click()
Text1.Text = Replace(Amin.mil2shams(Date$), "/", "-")
Frame1.Visible = True
End Sub
