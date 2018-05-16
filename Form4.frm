VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "äÍæå ãÍÇÓÈå ˜ÇÑãÒÏ"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "ãÏÊ ÒãÇä ÑÏÇÎÊ (˜á)"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ãÈáÛ ÈÇÞíãÇäÏå ÇÕá ÇÞÓÇØ"
      Height          =   495
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4920
      Width           =   10695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ãÈáÛ ÇÕá æÇã Èå åãÑÇå ˜ÇÑãÒÏ"
      Height          =   495
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ÊÚÏÇÏ ˜á ÇÞÓÇØ"
      Height          =   495
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ÔãÇÑå ÞÓØ"
      Height          =   495
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":2CFA
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16770252
      HeadLines       =   1
      RowHeight       =   28
      FormatLocked    =   -1  'True
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
         Caption         =   "˜Ï ÝÑãæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "promp"
         Caption         =   "ãÊä ÝÑãæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "promp2"
         Caption         =   "ÊæÖíÍÇÊ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6120
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4169.764
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ãÏÊ ÒãÇä ÑÏÇÎÊ (ãÇå ÞÈá)"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "äÑÎ ÈåÑå"
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      MaxLength       =   150
      TabIndex        =   7
      Top             =   4320
      Width           =   7815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ËÈÊ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "æíÑÇíÔ"
      Height          =   495
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ÇÖÇÝå"
      Height          =   495
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ãÈáÛ ÈÇÞíãÇäÏå ãÓÇæí æÇã"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ãÈáÛ ÇÕá æÇã"
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4320
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(A3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(B4)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊæÖíÍÇÊ"
      Height          =   495
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(B3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(B2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(B1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(A4)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "(A3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(A2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(A1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ãÊä ÝÑãæá"
      Height          =   495
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "˜Ï ÝÑãæá :"
      Height          =   495
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   9120
      Left            =   0
      Picture         =   "Form4.frx":2D0F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11865
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(A2)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command10_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(B3)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command11_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(B4)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command12_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(A3)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command2_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(A1)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command3_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(A3)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command4_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(A4)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command5_Click()
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

Private Sub Command6_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  DataGrid1.Col = 0
  Label2.Caption = DataGrid1.Text
  DataGrid1.Col = 1
  Text1.Text = DataGrid1.Text
  DataGrid1.Col = 2
  Text2.Text = DataGrid1.Text
  Label4.Caption = 2
End If
End Sub

Private Sub Command7_Click()
If (Label2.Caption = "") Or (Text1.Text = "") Then
  MsgBox "áØÝÇ ÊãÇãí ÝíáÏ åÇ ÑÇ Ê˜ãíá äãÇííÏ", vbCritical + vbMsgBoxRight, ""
  Exit Sub
End If
If Label4.Caption = 1 Then
  Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!ID = Label2.Caption
  Adodc1.Recordset.Fields!promp = Text1.Text
  Adodc1.Recordset.Fields!promp2 = Text2.Text
  Adodc1.Recordset.Update
  MsgBox "ÇØáÇÚÇÊ ÈÇ ãæÝÞíÊ ËÈÊ ÔÏ", vbInformation + vbMsgBoxRight, ""
End If
If Label4.Caption = 2 Then
  Adodc1.Refresh
  Adodc1.Recordset.Find "id=" + Label2.Caption, , adSearchForward, 1
  Adodc1.Recordset.Fields!promp = Text1.Text
  Adodc1.Recordset.Fields!promp2 = Text2.Text
  Adodc1.Recordset.Update
  MsgBox "ÇØáÇÚÇÊ ÈÇ ãæÝÞíÊ ÊÛííÑ íÏÇ ˜ÑÏ", vbInformation + vbMsgBoxRight, ""
End If
DataGrid1.Refresh
Command5.SetFocus
End Sub

Private Sub Command8_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(B1)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Command9_Click()
Text1.Text = Left(Text1.Text, Text1.SelStart) + "(B2)" + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
End Sub

Private Sub Form_Activate()
Form4.Adodc1.ConnectionString = Form1.Text1.Text
Form4.Adodc1.CommandType = adCmdUnknown
Form4.Adodc1.RecordSource = "SELECT * FROM formol ORDER BY id ASC"
Form4.Adodc1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 40) And (KeyAscii <= 57)) Or (KeyAscii = 8) Or (KeyAscii = 94) Then
Else
  KeyAscii = 0
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command7.SetFocus
End Sub

