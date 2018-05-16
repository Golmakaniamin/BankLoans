VERSION 5.00
Object = "{DF1D4B1E-D56E-4A40-BA98-2CC06080E796}#1.0#0"; "Tiny.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„œÌ—Ì  «⁄ »«—« "
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "ê“«—‘«  ÃœÌœ"
      Height          =   495
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   480
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TINYLib.Tiny Tiny1 
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "œ—»«—Â „«"
      Height          =   495
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ‰ŸÌ„« "
      Height          =   495
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ê÷⁄Ì "
      Height          =   495
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ê“«—‘« "
      Height          =   495
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ⁄«—Ì›"
      Height          =   495
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«‰ Œ«» ê—ÊÂ :"
      Height          =   495
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   9135
      Left            =   0
      Picture         =   "Form1.frx":2CFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoConnection1 As ADODB.Connection
Dim cmd1 As ADODB.Command
Dim adoRecordset1 As ADODB.Recordset
Dim tin1 As String
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Combo1_Click()
Form5.Adodc1.ConnectionString = Text1.Text
Form5.Adodc1.CommandType = adCmdUnknown
Form5.Adodc1.RecordSource = "SELECT * FROM Bank WHERE Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + " ORDER BY id"
Form5.Adodc1.Refresh
End Sub

Private Sub Command5_Click()
'Groupme
Form16.Show
End Sub

Private Sub Command6_Click()
Form20.Show
End Sub

Private Sub Form_Activate()
'Text2.Text = "„Ã „⁄ ‰”«ÃÌ —‰êÌ‰"
Text2.Text = "‘—ò   Ê·ÌœÌ ”Ì„ Ê ò«»· ”„‰«‰"
End Sub

Private Sub Form_Initialize()
 InitCommonControls
End Sub

Private Sub Command1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Form10.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Form6.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Form7.Show
Me.Hide
End Sub

Private Sub Form_Load()
'filenames$ = App.Path & "\MeLo.Exe"
'Open filenames$ For Input As #1
'Do While Not EOF(1)
'  Input #1, w
'Loop
'
'Close #1
'If w = "ok" Then
'  Tiny1.ServerIP = "192.168.160.1"
'  Tiny1.NetWorkINIT = True
''  Tiny1.Initialize = True
'  If Tiny1.TinyErrCode = 0 Then
'    Tiny1.UserPassWord = "5A85A8476F8F8FFB258A3FE48D162"
'    Tiny1.SpecialID = "v25f192510******"
'    Tiny1.ShowTinyInfo = True
'    DoEvents
'    If Tiny1.TinyErrCode = 0 Then
'      If Tiny1.DataPartition = "PraticGroup" Then
'        DoEvents
'        If Tiny1.SerialNumber = "2009-8803-1048" Then
'          DoEvents
'        Else
'          filenames$ = App.Path & "\MeLo.Exe"
'          Open filenames$ For Output As #1
'            Print #1, "Cancel"
'          Close #1
'          MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'          End
'        End If
'      Else
'        filenames$ = App.Path & "\MeLo.Exe"
'        Open filenames$ For Output As #1
'          Print #1, "Cancel"
'        Close #1
'        MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'        End
'      End If
'    Else
'      MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'      End
'    End If
'    Tiny1.ShowTinyInfo = False
'  Else
'    If Tiny1.TinyErrCode = 1 Then
'      MsgBox "⁄œ„ ‘‰«”«ÌÌ ﬁ›·", vbCritical + vbMsgBoxRight, ""
'    End If
'    End
'  End If
'Else
'  MsgBox "·ÿ›« »« Å‘ Ì»«‰Ì ‰—„ «›“«—  „«” »êÌ—Ìœ" + Chr(10) + Chr(13) + "09192595284", vbCritical + vbMsgBoxRight, ""
'  End
'End If

'”„‰«‰
Text1.Text = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Account_vam"

'¬—œ
'Text1.Text = "Provider=SQLOLEDB.1;Password=111111;Persist Security Info=True;User ID=Pratic_vam;Initial Catalog=Account_vam;Data Source=SERVER-2003"

'‰êÌ‰
'Text1.Text = "Provider=SQLOLEDB.1;Password=111111;Persist Security Info=True;User ID=Pratic_vam;Initial Catalog=Account_vam;Data Source=SERVER"

'«„Ì‰
'Text1.Text = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Account_vam"

Form17.Adodc1.ConnectionString = Text1.Text
Form17.Adodc1.CommandType = adCmdUnknown
Form17.Adodc1.RecordSource = "SELECT * FROM allGroup"
Form17.Adodc1.Refresh
  
Form17.Adodc1.Recordset.MoveFirst
Do
   Combo1.AddItem Form17.Adodc1.Recordset.Fields!Name
   Combo2.AddItem Form17.Adodc1.Recordset.Fields!ID
   Form17.Adodc1.Recordset.MoveNext
Loop Until Form17.Adodc1.Recordset.EOF = True
Combo1.ListIndex = 0

Form5.Adodc1.ConnectionString = Text1.Text
Form5.Adodc1.CommandType = adCmdUnknown
Form5.Adodc1.RecordSource = "SELECT * FROM Bank WHERE Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + " ORDER BY id"
Form5.Adodc1.Refresh

Form2.Adodc1.ConnectionString = Text1.Text
Form2.Adodc1.CommandType = adCmdUnknown
Form2.Adodc1.RecordSource = "SELECT * FROM vamtable WHERE rad=0"
Form2.Adodc1.Refresh

Form3.Adodc1.ConnectionString = Text1.Text
Form3.Adodc1.CommandType = adCmdUnknown
Form3.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE rad=0"
Form3.Adodc1.Refresh

Form4.Adodc1.ConnectionString = Text1.Text
Form4.Adodc1.CommandType = adCmdUnknown
Form4.Adodc1.RecordSource = "SELECT * FROM formol ORDER BY id ASC"
Form4.Adodc1.Refresh

Form6.Adodc1.ConnectionString = Text1.Text
Form6.Adodc1.CommandType = adCmdUnknown
Form6.Adodc1.RecordSource = "SELECT * FROM vamtable WHERE rad=0"
Form6.Adodc1.Refresh

Form7.Adodc1.ConnectionString = Text1.Text
Form7.Adodc1.CommandType = adCmdUnknown
Form7.Adodc1.RecordSource = "SELECT * FROM printvaz"
Form7.Adodc1.Refresh

Form7.Adodc2.ConnectionString = Text1.Text
Form7.Adodc2.CommandType = adCmdUnknown
Form7.Adodc2.RecordSource = "SELECT * FROM printtaahod"
Form7.Adodc2.Refresh

Form7.Adodc3.ConnectionString = Text1.Text
Form7.Adodc3.CommandType = adCmdUnknown
Form7.Adodc3.RecordSource = "SELECT * FROM Printbanksmoon"
Form7.Adodc3.Refresh

Form7.Adodc4.ConnectionString = Text1.Text
Form7.Adodc4.CommandType = adCmdUnknown
Form7.Adodc4.RecordSource = "SELECT * FROM vamtable WHERE idbank=0"
Form7.Adodc4.Refresh

Form9.Adodc1.ConnectionString = Text1.Text
Form9.Adodc1.CommandType = adCmdUnknown
Form9.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE rad=0"
Form9.Adodc1.Refresh
  
Form9.Adodc2.ConnectionString = Text1.Text
Form9.Adodc2.CommandType = adCmdUnknown
Form9.Adodc2.RecordSource = "SELECT * FROM aghsat WHERE rad=0"
Form9.Adodc2.Refresh
  
Form8.Adodc1.ConnectionString = Text1.Text
Form8.Adodc1.CommandType = adCmdUnknown
Form8.Adodc1.RecordSource = "SELECT * FROM printvazriz"
Form8.Adodc1.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
tin1 = Rnd
If Val(Right(Trim(Str(Round(tin1, 2))), 2)) = "35" Then
  DoEvents
'    Tiny1.UserPassWord = "5A85A8476F8F8FFB258A3FE48D162"
'    Tiny1.SpecialID = "v25f192510******"
'    Tiny1.ShowTinyInfo = True
'    DoEvents
'    If Tiny1.TinyErrCode = 0 Then
'      If Tiny1.DataPartition = "PraticGroup" Then
'        DoEvents
'        If Tiny1.SerialNumber = "2009-8803-1048" Then
'          DoEvents
'        Else
'          filenames$ = App.Path & "\MeLo.Exe"
'          Open filenames$ For Output As #1
'            Print #1, "Cancel"
'          Close #1
'          MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'          End
'        End If
'      Else
'        filenames$ = App.Path & "\MeLo.Exe"
'        Open filenames$ For Output As #1
'          Print #1, "Cancel"
'        Close #1
'        MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'        End
'      End If
'    Else
'      MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'      End
'    End If
'    Tiny1.ShowTinyInfo = False
End If
End Sub

Private Sub Tiny1_TinyDisconnect()
MsgBox 1
End Sub

Private Sub Tiny1_TinyHIDDidconnect()
MsgBox 2
End Sub
