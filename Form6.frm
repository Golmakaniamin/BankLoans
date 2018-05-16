VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê÷⁄Ì  Ê«„ Â«"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE4CC&
      Caption         =   " ”ÊÌÂ"
      Height          =   855
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2400
      Width           =   2775
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFE4CC&
         Caption         =   "Â„Â"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFE4CC&
         Caption         =   "‰‘œÂ"
         Height          =   375
         Left            =   960
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFE4CC&
         Caption         =   "‘œÂ"
         Height          =   375
         Left            =   1920
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   465
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Combo3"
      Top             =   2760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      ItemData        =   "Form6.frx":2CFA
      Left            =   11400
      List            =   "Form6.frx":2D1F
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   465
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   120
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":2D87
      Height          =   5655
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9975
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
            ColumnWidth     =   975.118
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   120
      Top             =   480
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ã„⁄ »œÂÌ :"
      Height          =   495
      Index           =   1
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   12
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "”ﬁ› Ê«„ :"
      Height          =   495
      Index           =   13
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ã„⁄ Ê«„ Â«Ì œ—Ì«› Ì :"
      Height          =   495
      Index           =   14
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   15
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   16
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "”ﬁ› »«“ :"
      Height          =   495
      Index           =   17
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ »«‰ò"
      Height          =   495
      Index           =   0
      Left            =   14400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "Form6.frx":2D9C
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset

Private Sub Combo1_Click()
If Combo1.ListIndex <> -1 Then
Adodc2.ConnectionString = Form1.Text1.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + ") and (tasvie='‰‘œÂ') ORDER BY rad ASC"
Adodc2.Refresh

www = Combo1.ListIndex

Label1(16).Caption = 0
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc2.Recordset.MoveFirst
  Do
      Form9.Adodc1.ConnectionString = Form1.Text1.Text
      Form9.Adodc1.CommandType = adCmdUnknown
      Form9.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank = " + Trim(Str(Adodc2.Recordset.Fields!idbank)) + ") AND (rad = " + Trim(Str(Adodc2.Recordset.Fields!rad)) + ") and (yesno='1') ORDER BY rad2 ASC"
      Form9.Adodc1.Refresh
      If Form9.Adodc1.Recordset.RecordCount > 0 Then
        Form9.Adodc1.Recordset.MoveLast
        Label1(16).Caption = Val(Label1(16).Caption) + Form9.Adodc1.Recordset.Fields!sumbedehiasl
      Else
        Label1(16).Caption = Val(Label1(16).Caption) + Adodc2.Recordset.Fields!moneyvamrial
      End If
    Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
End If

Label1(12).Caption = Combo3.List(Combo1.ListIndex)
Label1(15).Caption = ""
If Adodc2.Recordset.RecordCount > 0 Then
  Adodc2.Recordset.MoveFirst
  Do
    Label1(15).Caption = Val(Label1(15).Caption) + Val(Adodc2.Recordset.Fields!moneyvamrial)
    Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
End If

Label1(2).Caption = Val(Label1(12).Caption) - Val(Label1(16).Caption)

Label1(15).Caption = Amin.moneyaminjoda(Label1(15).Caption)
Label1(16).Caption = Amin.moneyaminjoda(Label1(16).Caption)
Label1(12).Caption = Amin.moneyaminjoda(Label1(12).Caption)
Label1(2).Caption = Amin.moneyaminjoda(Label1(2).Caption)

Adodc1.ConnectionString = Form1.Text1.Text
Adodc1.CommandType = adCmdUnknown
If Option1.Value = True Then Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + ") and (tasvie='‘œÂ') ORDER BY rad ASC"
If Option2.Value = True Then Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + ") and (tasvie='‰‘œÂ') ORDER BY rad ASC"
If Option3.Value = True Then Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo2.List(Combo1.ListIndex))) + ") ORDER BY rad ASC"
Adodc1.Refresh
End If
End Sub

Private Sub DataGrid1_DblClick()
If Adodc1.Recordset.RecordCount > 0 Then
  Form9.Label1(2).Caption = Adodc1.Recordset.Fields!rad
  Form9.Label1(4).Caption = Adodc1.Recordset.Fields!idbank
  Form9.Label1(0).Caption = Combo1.Text
  Form9.Adodc1.ConnectionString = Form1.Text1.Text
  Form9.Adodc1.CommandType = adCmdUnknown
  Form9.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") ORDER BY rad2"
  Form9.Adodc1.Refresh
  
  Form9.Adodc2.ConnectionString = Form9.Adodc1.ConnectionString
  Form9.Adodc2.CommandType = adCmdUnknown
  Form9.Adodc2.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!rad)) + ") and (yesno='1') ORDER BY rad2"
  Form9.Adodc2.Refresh
  If Form9.Adodc1.Recordset.RecordCount > 0 Then
    Form9.Adodc1.Recordset.MoveLast
    If Form9.Adodc1.Recordset.Fields!darsad = 0 Then Form9.DataGrid1.Columns.Item(1).Visible = False: Form9.DataGrid1.Columns.Item(2).Visible = True
    If Form9.Adodc1.Recordset.Fields!Money = 0 Then Form9.DataGrid1.Columns.Item(2).Visible = False: Form9.DataGrid1.Columns.Item(1).Visible = True
  
    db1.Open Form1.Text1.Text
      rs1.Open "SELECT * FROM formol WHERE (id=" + Adodc1.Recordset.Fields!idformol + " )", db1
      Form9.Label1(5).Caption = rs1.Fields!promp
    db1.Close
    Form9.Label2.Caption = 1
    Form9.Show
    Me.Hide
  End If
End If
End Sub

Private Sub Form_Load()
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
Me.Hide
End Sub

Private Sub Option1_Click()
Call Combo1_Click
End Sub

Private Sub Option2_Click()
Call Combo1_Click
End Sub

Private Sub Option3_Click()
Call Combo1_Click
End Sub

