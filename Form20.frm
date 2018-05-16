VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form20 
   BackColor       =   &H00FFE4CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ê“«—‘«  ÃœÌœ"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12525
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   12525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "—Ì“ »œÂÌ  «—ÌŒ"
      Height          =   495
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE4CC&
      Caption         =   "—Ì“ »œÂÌ"
      Height          =   3615
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "‰„«Ì‘"
         Height          =   495
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3000
         Width           =   2775
      End
      Begin VB.ComboBox Combo4 
         Height          =   465
         ItemData        =   "Form20.frx":0000
         Left            =   360
         List            =   "Form20.frx":000A
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         ItemData        =   "Form20.frx":0024
         Left            =   360
         List            =   "Form20.frx":0049
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   465
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„— » ”«“Ì"
         Height          =   495
         Index           =   19
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«“  «—ÌŒ"
         Height          =   495
         Index           =   2
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «  «—ÌŒ"
         Height          =   495
         Index           =   5
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ »«‰ò"
         Height          =   495
         Index           =   0
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "—Ì“ »œÂÌ"
      Height          =   495
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFE4CC&
      Caption         =   "—Ì“ »œÂÌ  «—ÌŒ"
      Height          =   3615
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox Combo6 
         Height          =   465
         ItemData        =   "Form20.frx":00B1
         Left            =   360
         List            =   "Form20.frx":00BB
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         Height          =   465
         ItemData        =   "Form20.frx":00D5
         Left            =   480
         List            =   "Form20.frx":00D7
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Text            =   "Combo5"
         Top             =   360
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox Combo3 
         Height          =   465
         ItemData        =   "Form20.frx":00D9
         Left            =   360
         List            =   "Form20.frx":00E3
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "‰„«Ì‘"
         Height          =   495
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3000
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ »«‰ò"
         Height          =   495
         Index           =   6
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «  «—ÌŒ"
         Height          =   495
         Index           =   4
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«“  «—ÌŒ"
         Height          =   495
         Index           =   3
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„— » ”«“Ì"
         Height          =   495
         Index           =   1
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2400
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub amin_hide()
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim rs(2) As New ADODB.Recordset

Combo1.Clear
Combo2.Clear

db1.Open Form1.Text1.Text
  rs(0).Open "SELECT * FROM Bank Where (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") ORDER BY id", db1
    rs(0).MoveFirst
    Do
      Combo2.AddItem rs(0).Fields!ID
      Combo1.AddItem rs(0).Fields!Name
      rs(0).MoveNext
    Loop Until rs(0).EOF = True
  rs(0).Close
db1.Close
Combo1.ListIndex = 0

Call amin_hide
Frame1.Visible = True
End Sub

Private Sub Command2_Click()
Dim amin_form21_8 As New Form21
amin_form21_8.Label1.Caption = 8
amin_form21_8.Show
End Sub

Private Sub Command3_Click()
Dim db1 As New ADODB.Connection
Dim rs(2) As New ADODB.Recordset

Combo3.Clear
Combo5.Clear

db1.Open Form1.Text1.Text
  rs(0).Open "SELECT * FROM Bank Where (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") ORDER BY id", db1
    rs(0).MoveFirst
    Do
      Combo5.AddItem rs(0).Fields!ID
      Combo3.AddItem rs(0).Fields!Name
      rs(0).MoveNext
    Loop Until rs(0).EOF = True
  rs(0).Close
db1.Close
Combo1.ListIndex = 0

Call amin_hide
Frame2.Visible = True
End Sub

Private Sub Command4_Click()
Dim amin_form21_9 As New Form21
amin_form21_9.Label1.Caption = 9
amin_form21_9.Show
End Sub

Private Sub Form_Load()
MaskEdBox1.Mask = "13##/##/##"
MaskEdBox1.PromptChar = "*"

MaskEdBox2.Mask = "13##/##/##"
MaskEdBox2.PromptChar = "*"

MaskEdBox3.Mask = "13##/##/##"
MaskEdBox3.PromptChar = "*"

MaskEdBox4.Mask = "13##/##/##"
MaskEdBox4.PromptChar = "*"

End Sub

