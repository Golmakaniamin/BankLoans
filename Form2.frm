VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄—Ì› Ê«„"
   ClientHeight    =   11220
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11220
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   " »œÌ· »Â  ”ÊÌÂ ‰‘œÂ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   7
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   465
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Text            =   "Combo3"
      Top             =   2400
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   6
      Left            =   1560
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4200
      Width           =   11655
   End
   Begin VB.ComboBox Combo4 
      Height          =   465
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   4
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "À» "
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Õ–›"
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÊÌ—«Ì‘"
      Height          =   495
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«÷«›Â"
      Height          =   495
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":2CFA
      Height          =   5535
      Left            =   240
      TabIndex        =   16
      Top             =   5520
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16770252
      HeadLines       =   1
      RowHeight       =   29
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "rad"
         Caption         =   "—œÌ›"
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
         DataField       =   "vamcode"
         Caption         =   "ﬂœ Ê«„"
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
         DataField       =   "novam"
         Caption         =   "‰Ê⁄ Ê«„"
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
      BeginProperty Column03 
         DataField       =   "dateshamsi"
         Caption         =   " «—ÌŒ œ—Ì«›  ‘„”Ì"
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
      BeginProperty Column04 
         DataField       =   "datemiladi"
         Caption         =   " «—ÌŒ œ—Ì«›  „Ì·«œÌ"
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
      BeginProperty Column05 
         DataField       =   "moneyvamarz"
         Caption         =   "„»·€ Ê«„ «—“Ì"
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
      BeginProperty Column06 
         DataField       =   "nerkharz"
         Caption         =   "‰—Œ «—“"
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
      BeginProperty Column07 
         DataField       =   "noarz"
         Caption         =   "‰Ê⁄ «—“"
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
      BeginProperty Column08 
         DataField       =   "moneyvamrial"
         Caption         =   "„»·€ Ê«„ —Ì«·Ì"
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
      BeginProperty Column09 
         DataField       =   "allbahre"
         Caption         =   "„»·€ ò«—„“œ ﬂ· Ê«„ —Ì«·Ì"
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
      BeginProperty Column10 
         DataField       =   "idformol"
         Caption         =   "‘„«—Â ›—„Ê·"
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
      BeginProperty Column11 
         DataField       =   "subject"
         Caption         =   "„Ê÷Ê⁄"
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
      BeginProperty Column12 
         DataField       =   "tasvie"
         Caption         =   " ”ÊÌÂ"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   0
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Text            =   "Combo2"
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      ItemData        =   "Form2.frx":2D0F
      Left            =   11520
      List            =   "Form2.frx":2D34
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   5
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   3
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   2
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "1"
      Top             =   3000
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Index           =   1
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2520
      Top             =   4080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1200
      Top             =   4080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂœ ›—„Ê· ò«—„“œ"
      Height          =   495
      Index           =   12
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «—“"
      Height          =   495
      Index           =   4
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ Ê«„"
      Height          =   495
      Index           =   19
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ê÷Ê⁄"
      Height          =   495
      Index           =   18
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ Ê«„ «—“Ì"
      Height          =   495
      Index           =   11
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   10
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ› :"
      Height          =   495
      Index           =   6
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ Ê«„ —Ì«·Ì"
      Height          =   495
      Index           =   8
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰—Œ «—“"
      Height          =   495
      Index           =   7
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ œ—Ì«›  „Ì·«œÌ"
      Height          =   495
      Index           =   5
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ò«—„“œﬂ· Ê«„ —Ì«·Ì"
      Height          =   495
      Index           =   3
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ œ—Ì«›  ‘„”Ì"
      Height          =   495
      Index           =   2
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ Ê«„"
      Height          =   495
      Index           =   1
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ »«‰ò"
      Height          =   495
      Index           =   0
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "Form2.frx":2D9C
      Top             =   -120
      Width           =   15360
   End
   Begin VB.Menu mnutarif 
      Caption         =   " ⁄«—Ì›"
      Begin VB.Menu mnubanks 
         Caption         =   " ⁄—Ì› »«‰ò"
      End
      Begin VB.Menu mnuformol 
         Caption         =   "›—„Ê· »Â—Â"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblclickoc As Boolean
Dim adoConnection1 As ADODB.Connection
Dim cmd1 As ADODB.Command
Dim adoRecordset1 As ADODB.Recordset
Dim www As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Sub new_1()
For q = 0 To 7
  Text1(q).Text = ""
Next q
End Sub

Private Sub Combo1_Click()
Adodc1.ConnectionString = Form1.Text1.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + ") ORDER BY rad ASC"
Adodc1.Refresh
End Sub

Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox1.SetFocus
End Sub

Private Sub Combo4_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
Dim tmp
tmp = SendMessage(Combo4.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub Command1_Click()
Call new_1
If Combo1.ListIndex <> -1 Then
  Label2.Caption = 1
  q = 1
  If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveFirst
    Do
      If Adodc1.Recordset.Fields!rad <> q Then Exit Do
      q = q + 1
      Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
  End If
  Label1(10).Caption = q
  Text1(0).SetFocus

  Adodc2.ConnectionString = Form1.Text1.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "SELECT DISTINCT novam FROM vamtable"
  Adodc2.Refresh
  Combo4.Clear
  If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveFirst
    Do
      Combo4.AddItem Adodc2.Recordset.Fields!novam
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
  End If
  Adodc2.ConnectionString = Form1.Text1.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "SELECT * FROM vamtable"
  Adodc2.Refresh
End If
End Sub

Private Sub Command2_Click()
Call new_1
If Combo1.ListIndex <> -1 Then
  If Adodc1.Recordset.RecordCount > 0 Then
    DataGrid1.Col = 0
    Label1(10).Caption = DataGrid1.Text
    
    DataGrid1.Col = 1
    Text1(0).Text = DataGrid1.Text
    
    DataGrid1.Col = 2
    Combo4.Text = DataGrid1.Text
    
    DataGrid1.Col = 3
    MaskEdBox1.Text = DataGrid1.Text
    
    DataGrid1.Col = 4
    MaskEdBox2.Text = DataGrid1.Text
    
    DataGrid1.Col = 5
    Text1(1).Text = DataGrid1.Text
    
    DataGrid1.Col = 6
    Text1(2).Text = DataGrid1.Text
    
    DataGrid1.Col = 7
    Text1(3).Text = DataGrid1.Text
    
    DataGrid1.Col = 8
    Text1(7).Text = DataGrid1.Text
    
    DataGrid1.Col = 9
    Text1(4).Text = DataGrid1.Text
    
    DataGrid1.Col = 10
    Text1(5).Text = DataGrid1.Text
    
    DataGrid1.Col = 11
    Text1(6).Text = DataGrid1.Text
    
    Label2.Caption = 2
  End If
End If
End Sub

Private Sub Command3_Click()
e = MsgBox("¬Ì« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
If (Adodc1.Recordset.RecordCount > 0) And (e = 6) Then
  Set adoConnection1 = New ADODB.Connection
  adoConnection1.ConnectionString = Form1.Text1.Text
  adoConnection1.CursorLocation = adUseClient
  adoConnection1.Open
  Set cmd1 = New ADODB.Command
  cmd1.ActiveConnection = adoConnection1
  cmd1.CommandText = "DELETE FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Form2.Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Form2.Adodc1.Recordset.Fields!rad)) + ")"
  cmd1.CommandType = adCmdText
  cmd1.Properties.Refresh
  Set adoRecordset1 = cmd1.Execute
  
  Adodc1.Recordset.Delete
  Adodc1.Refresh
  DataGrid1.Refresh
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + ") ORDER BY rad ASC"
  Adodc1.Refresh
  DataGrid1.Refresh
End If
End Sub

Private Sub Command4_Click()

If Label1(10).Caption = "" Then
  e = MsgBox("·ÿ›« œò„Â «÷«›Â —« ò·Ìò ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Command1.SetFocus
  Exit Sub
End If

If Text1(0).Text = "" Then
  e = MsgBox("›Ì·œ ﬂœ Ê«„ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(0).SetFocus
  Exit Sub
End If

'If Combo4.ListIndex = -1 Then
'  e = MsgBox("›Ì·œ ‰Ê⁄ Ê«„ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ÌòÌ «“ Õ«·«  ÅÌ‘ ›—÷ —« «‰ Œ«» ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
'  Combo4.SetFocus
'  Exit Sub
'End If

If MaskEdBox1.Text = "" Then
  e = MsgBox("›Ì·œ  «—ÌŒ œ—Ì«›  ‘„”Ì Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If InStr(MaskEdBox1.Text, "*") <> 0 Then
  e = MsgBox("·ÿ›« ›Ì·œ  «—ÌŒ œ—Ì«›  ‘„”Ì —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If MaskEdBox2.Text = "" Then
  e = MsgBox("›Ì·œ  «—ÌŒ œ—Ì«›  „Ì·«œÌ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If InStr(MaskEdBox2.Text, "*") <> 0 Then
  e = MsgBox("·ÿ›« ›Ì·œ  «—ÌŒ œ—Ì«›  „Ì·«œÌ —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If Text1(1).Text = "" Then
  e = MsgBox("›Ì·œ „»·€ Ê«„ «—“Ì Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(1).SetFocus
  Exit Sub
End If

If Text1(2).Text = "" Then
  e = MsgBox("›Ì·œ ‰—Œ «—“ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(2).SetFocus
  Exit Sub
End If

If Text1(3).Text = "" Then
  e = MsgBox("›Ì·œ ‰Ê⁄ «—“ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(3).SetFocus
  Exit Sub
End If

If Text1(7).Text = "" Then
  e = MsgBox("›Ì·œ „»·€ Ê«„ —Ì«·Ì Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(7).SetFocus
  Exit Sub
End If

If Text1(4).Text = "" Then
  e = MsgBox("›Ì·œ „»·€ ò«—„“œ ﬂ· Ê«„ —Ì«·Ì Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(4).SetFocus
  Exit Sub
End If

If Text1(5).Text = "" Then
  e = MsgBox("›Ì·œ ﬂœ ›—„Ê· ò«—„“œ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(5).SetFocus
  Exit Sub
End If

If Text1(6).Text = "" Then
  e = MsgBox("›Ì·œ „Ê÷Ê⁄ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1(6).SetFocus
  Exit Sub
End If

If Label2.Caption = 1 Then
  Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!idbank = Combo2.List(Combo1.ListIndex)
  Adodc1.Recordset.Fields!rad = Label1(10).Caption
  Adodc1.Recordset.Fields!vamcode = Text1(0).Text
  Adodc1.Recordset.Fields!Subject = Text1(6).Text
  Adodc1.Recordset.Fields!novam = Combo4.Text
  Adodc1.Recordset.Fields!dateshamsi = MaskEdBox1.Text
  Adodc1.Recordset.Fields!datemiladi = MaskEdBox2.Text
  Adodc1.Recordset.Fields!moneyvamarz = Text1(1).Text
  Adodc1.Recordset.Fields!nerkharz = Text1(2).Text
  Adodc1.Recordset.Fields!noarz = Text1(3).Text
  Adodc1.Recordset.Fields!moneyvamrial = Text1(7).Text
  Adodc1.Recordset.Fields!allbahre = Text1(4).Text
  Adodc1.Recordset.Fields!idformol = Text1(5).Text
  Adodc1.Recordset.Fields!Groupme = Form1.Combo2.List(Form1.Combo1.ListIndex)
  Adodc1.Recordset.Fields!tasvie = "‰‘œÂ"
  Adodc1.Recordset.Update
  MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  À»  ‘œ", vbInformation + vbMsgBoxRight, ""
End If
If Label2.Caption = 2 Then
  Adodc1.Refresh
  Adodc1.Recordset.Find "rad=" + Label1(10).Caption, , adSearchForward, 1
  Adodc1.Recordset.Fields!idbank = Combo2.List(Combo1.ListIndex)
  Adodc1.Recordset.Fields!vamcode = Text1(0).Text
  Adodc1.Recordset.Fields!Subject = Text1(6).Text
  Adodc1.Recordset.Fields!novam = Combo4.Text
  Adodc1.Recordset.Fields!dateshamsi = MaskEdBox1.Text
  Adodc1.Recordset.Fields!datemiladi = MaskEdBox2.Text
  Adodc1.Recordset.Fields!moneyvamarz = Text1(1).Text
  Adodc1.Recordset.Fields!nerkharz = Text1(2).Text
  Adodc1.Recordset.Fields!noarz = Text1(3).Text
  Adodc1.Recordset.Fields!moneyvamrial = Text1(7).Text
  Adodc1.Recordset.Fields!allbahre = Text1(4).Text
  Adodc1.Recordset.Fields!idformol = Text1(5).Text
  Adodc1.Recordset.Fields!Groupme = Form1.Combo2.List(Form1.Combo1.ListIndex)
  Adodc1.Recordset.Fields!tasvie = "‰‘œÂ"
  Adodc1.Recordset.Update
  MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ   €ÌÌ— ÅÌœ« ò—œ", vbInformation + vbMsgBoxRight, ""
End If

'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "SELECT * FROM vamtable WHERE idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + " ORDER BY rad ASC"
'Adodc1.Refresh
'DataGrid1.Refresh
'Label1(15).Caption = 0
'If Form2.Adodc1.Recordset.RecordCount > 0 Then
'  Form2.Adodc1.Recordset.MoveFirst
'  Do
'    Label1(15).Caption = Val(Label1(15).Caption) + Form2.Adodc1.Recordset.Fields!moneyvam
'    Form2.Adodc1.Recordset.MoveNext
'  Loop Until Form2.Adodc1.Recordset.EOF = True
'End If
'Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command5_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!tasvie = "‰‘œÂ"
  Adodc1.Recordset.Update
  MsgBox "⁄„·Ì«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", vbInformation, ""
End If
End Sub

Private Sub DataGrid1_DblClick()
If Adodc1.Recordset.RecordCount > 0 Then
  Form3.Label1(2).Caption = Adodc1.Recordset.Fields!rad
  Form3.Label1(4).Caption = Adodc1.Recordset.Fields!idbank
  Form3.Label1(0).Caption = Combo1.Text
    
  Form3.Adodc1.ConnectionString = Form1.Text1.Text
  Form3.Adodc1.CommandType = adCmdUnknown
  Form3.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") ORDER BY rad2"
  Form3.Adodc1.Refresh
  Form3.Show
  Me.Hide
End If
End Sub

Private Sub Form_Load()
MaskEdBox1.Mask = "13##/##/##"
MaskEdBox1.PromptChar = "*"
MaskEdBox2.Mask = "20##/##/##"
MaskEdBox2.PromptChar = "*"
Combo1.Clear
Combo2.Clear
Combo3.Clear
If Form5.Adodc1.Recordset.RecordCount > 0 Then
  Form5.Adodc1.Recordset.MoveFirst
  Do
    Combo2.AddItem Form5.Adodc1.Recordset.Fields!ID
    Combo1.AddItem Form5.Adodc1.Recordset.Fields!Name
    Combo3.AddItem Form5.Adodc1.Recordset.Fields!saghf
    Form5.Adodc1.Recordset.MoveNext
  Loop Until Form5.Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox2.SetFocus
End Sub

Private Sub MaskEdBox1_LostFocus()
If InStr(MaskEdBox1.Text, "*") = 0 Then
  MaskEdBox2.Text = Amin.shams2mil(MaskEdBox1.Text)
End If
End Sub

Private Sub MaskEdBox2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1(1).SetFocus
End Sub

Private Sub MaskEdBox2_LostFocus()
If InStr(MaskEdBox2.Text, "*") = 0 Then
  Dim aa1 As String
  aa1 = Mid(MaskEdBox2.Text, 6, 2) + "-" + Mid(MaskEdBox2.Text, 9, 2) + "-" + Mid(MaskEdBox2.Text, 1, 4)
  MaskEdBox1.Text = Amin.mil2shams(aa1)
End If
End Sub

Private Sub mnubanks_Click()
Form5.Show
End Sub

Private Sub mnuformol_Click()
Form4.Show
End Sub

Private Sub Text1_Change(Index As Integer)
If (Index = 1) Or (Index = 2) Then
  If (Val(Text1(1)) <> 0) Or (Val(Text1(2).Text) <> 0) Then
    Text1(7).Text = Val(Text1(1).Text) * Val(Text1(2).Text)
  End If
End If

If (Index = 7) Then
  If (Val(Text1(7)) <> 0) Or (Val(Text1(2).Text) <> 0) Then
    Text1(1).Text = Val(Text1(7).Text) / Val(Text1(2).Text)
  End If
End If
End Sub

Private Sub Text1_DblClick(Index As Integer)
Form4.Show
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Select Case Index
    Case 0
      Combo4.SetFocus
      
    Case 1
      Text1(2).SetFocus
  
    Case 2
      Text1(3).SetFocus
  
    Case 3
      Text1(7).SetFocus
  
    Case 7
      Text1(4).SetFocus
  
    Case 4
      Text1(5).SetFocus
  
    Case 5
      Text1(6).SetFocus
  
    Case 6
      Command4.SetFocus
  
  End Select
End If
End Sub
