VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFE4CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄—Ì› »«‰ò"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":2CFA
      Height          =   3255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16770252
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "òœ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "name"
         Caption         =   "‰«„"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "saghf"
         Caption         =   "”ﬁ›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1769.953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2640
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÊÌ—«Ì‘"
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "À» "
      Height          =   495
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”ﬁ›"
      Height          =   495
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ »«‰ò :"
      Height          =   495
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ »«‰ò "
      Height          =   495
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   255
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer

Private Sub Command1_Click()
Label4.Caption = 1
Text1.Text = ""
q = 1
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Sort = "id"
  Adodc1.Recordset.MoveFirst
  Do
    If Adodc1.Recordset.Fields!ID <> q Then Exit Do
    q = q + 1
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
Label2.Caption = q
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  DataGrid1.Col = 0
  Label2.Caption = DataGrid1.Text
  DataGrid1.Col = 1
  Text1.Text = DataGrid1.Text
  DataGrid1.Col = 2
  Text2.Text = DataGrid1.Text
'  DataGrid1.Col = 3
'  Text3.Text = DataGrid1.Text
  Label4.Caption = 2
End If
End Sub

Private Sub Command3_Click()
If (Label2.Caption = "") Or (Text1.Text = "") Then
  MsgBox "·ÿ›«  „«„Ì ›Ì·œ Â« —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, ""
  Exit Sub
End If
If Label4.Caption = 1 Then
  Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!ID = Label2.Caption
  Adodc1.Recordset.Fields!Name = Text1.Text
  Adodc1.Recordset.Fields!saghf = Text2.Text
  Adodc1.Recordset.Fields!Groupme = Form1.Combo2.List(Form1.Combo1.ListIndex)
  Adodc1.Recordset.Update
  MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation + vbMsgBoxRight, ""
End If
If Label4.Caption = 2 Then
  Adodc1.Refresh
  Adodc1.Recordset.Find "id=" + Label2.Caption, , adSearchForward, 1
  Adodc1.Recordset.Fields!Name = Text1.Text
  Adodc1.Recordset.Fields!saghf = Text2.Text
  Adodc1.Recordset.Fields!Groupme = Form1.Combo2.List(Form1.Combo1.ListIndex)
  Adodc1.Recordset.Update
  MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ   €ÌÌ— ÅÌœ« ò—œ", vbInformation + vbMsgBoxRight, ""
End If
DataGrid1.Refresh
Command1.SetFocus
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = Form1.Text1.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM Bank WHERE Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + " ORDER BY id"
Adodc1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command3.SetFocus
End Sub

