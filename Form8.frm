VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00FFE4CC&
   Caption         =   "ÑíÒ æÖÚíÊ ÊÓåíáÇÊ ÏÑíÇÝÊí ÇÒ ÈÇä˜"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ç"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   10440
      Width           =   15015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form8.frx":2CFA
      Height          =   10215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   18018
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "rad"
         Caption         =   "ÑÏíÝ"
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
         DataField       =   "idgarardad"
         Caption         =   "ÔãÇÑå ÞÑÇÑÏÇÏ"
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
         DataField       =   "AllVamAsl"
         Caption         =   "˜á ãÈáÛ æÇã ÑíÇáí"
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
      BeginProperty Column03 
         DataField       =   "AllVambahre"
         Caption         =   "˜ÇÑãÒÏ ˜á æÇã"
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
      BeginProperty Column04 
         DataField       =   "mablaghgestasl"
         Caption         =   "ãÈáÛ ÞÓØ - ÇÕá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "mablaghgestbahre"
         Caption         =   "ãÈáÛ ÞÓØ - ˜ÇÑãÒÏ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "sumpardakhtasl"
         Caption         =   "ãÈáÛ ÑÏÇÎÊí - ÇÕá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "sumpardakhtbahre"
         Caption         =   "ãÈáÛ ÑÏÇÎÊí - ˜ÇÑãÒÏ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "sumbedehiasl"
         Caption         =   "ãÈáÛ ÈÏåí - ÇÕá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "sumbedehibahre"
         Caption         =   "ãÈáÛ ÈÏåí - ˜ÇÑãÒÏ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "mablaghgestall"
         Caption         =   "ãÈáÛ ÞÓØ - ÌãÚ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "sumpardakhtall"
         Caption         =   "ãÈáÛ ÑÏÇÎÊí - ÌãÚ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "sumbedehiall"
         Caption         =   "ãÈáÛ ÈÏåí - ÌãÚ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ÑíÇá"" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "numberall"
         Caption         =   "ÊÚÏÇÏ ˜á ÇÞÓÇØ"
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
      BeginProperty Column14 
         DataField       =   "numberdif"
         Caption         =   "ÊÚÏÇÏ ÇÞÓÇØ ÑÏÇÎÊ ÔÏå"
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
      BeginProperty Column15 
         DataField       =   "numbermonde"
         Caption         =   "ÊÚÏÇÏ ÇÞÓÇØ ÈÇÞí ãÇäÏå"
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
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1964.976
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1874.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   240
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs(2) As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub Command1_Click()
Adodc1.Refresh
Adodc1.Refresh
Form12.Show
End Sub

Private Sub DataGrid1_DblClick()
If Adodc1.Recordset.RecordCount > 0 Then
  Form9.Label1(2).Caption = Adodc1.Recordset.Fields!idrad
  Form9.Label1(4).Caption = Adodc1.Recordset.Fields!idbank
  db1.Open Form1.Text1.Text
    rs1.Open "SELECT * FROM Bank WHERE (id=" + Adodc1.Recordset.Fields!idbank + ")", db1
      Form9.Label1(0).Caption = rs1.Fields!Name
    rs1.Close
      
    Form9.Adodc1.ConnectionString = Form1.Text1.Text
    Form9.Adodc1.CommandType = adCmdUnknown
    Form9.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!idrad)) + ") ORDER BY rad2"
    Form9.Adodc1.Refresh
  
    Form6.Adodc1.ConnectionString = Form1.Text1.Text
    Form6.Adodc1.CommandType = adCmdUnknown
    Form6.Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!idrad)) + ")"
    Form6.Adodc1.Refresh
  
    Form9.Adodc2.ConnectionString = Form9.Adodc1.ConnectionString
    Form9.Adodc2.CommandType = adCmdUnknown
    Form9.Adodc2.RecordSource = "SELECT * FROM aghsat WHERE (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!idrad)) + ") and (yesno='1') ORDER BY rad2"
    Form9.Adodc2.Refresh
    
    If Form9.Adodc1.Recordset.RecordCount > 0 Then
      Form9.Adodc1.Recordset.MoveLast
      If Form9.Adodc1.Recordset.Fields!darsad = 0 Then Form9.DataGrid1.Columns.Item(1).Visible = False: Form9.DataGrid1.Columns.Item(2).Visible = True
      If Form9.Adodc1.Recordset.Fields!Money = 0 Then Form9.DataGrid1.Columns.Item(2).Visible = False: Form9.DataGrid1.Columns.Item(1).Visible = True
  
      rs2.Open "SELECT * FROM vamtable WHERE (idbank=" + Trim(Str(Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Adodc1.Recordset.Fields!idrad)) + ")", db1
        rs1.Open "SELECT * FROM formol WHERE (id=" + rs2.Fields!idformol + " )", db1
          Form9.Label1(5).Caption = rs1.Fields!promp
        rs1.Close
      rs2.Close
       Form9.Label2.Caption = 2
       Form9.Show
      Me.Hide
    End If
  db1.Close
End If
End Sub

Private Sub Form_Activate()
Dim info(13) As String
  db1.Open Form1.Text1.Text
    rs1.Open "DELETE FROM printvazriz", db1
  db1.Close
  
  Adodc1.ConnectionString = Form1.Text1.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM printvazriz ORDER BY rad ASC"
  Adodc1.Refresh
  
  info(12) = 1
  For q = 0 To 11
    info(q) = 0
  Next q
  Form2.Adodc1.ConnectionString = Form1.Text1.Text
  Form2.Adodc1.CommandType = adCmdUnknown
  Form2.Adodc1.RecordSource = "SELECT * FROM vamtable WHERE (idbank=" + Text1.Text + ") and (tasvie='äÔÏå') ORDER BY rad ASC"
  Form2.Adodc1.Refresh
  If Form2.Adodc1.Recordset.RecordCount > 0 Then
    Form2.Adodc1.Recordset.MoveFirst
    Do
      Form9.Adodc1.ConnectionString = Form1.Text1.Text
      Form9.Adodc1.CommandType = adCmdUnknown
      Form9.Adodc1.RecordSource = "SELECT * FROM aghsat WHERE (idbank=" + Trim(Str(Form2.Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Form2.Adodc1.Recordset.Fields!rad)) + ") ORDER BY rad2"
      Form9.Adodc1.Refresh
      
      Form9.Adodc2.ConnectionString = Form1.Text1.Text
      Form9.Adodc2.CommandType = adCmdUnknown
      Form9.Adodc2.RecordSource = "SELECT * FROM aghsat WHERE (idbank=" + Trim(Str(Form2.Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Form2.Adodc1.Recordset.Fields!rad)) + ") and (yesno='1') ORDER BY rad2"
      Form9.Adodc2.Refresh
      
      If Form9.Adodc2.Recordset.RecordCount > 0 Then
        Form9.Adodc2.Recordset.MoveLast
        Adodc1.Refresh
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!idrad = Form2.Adodc1.Recordset.Fields!rad
        Adodc1.Recordset.Fields!idbank = Form2.Adodc1.Recordset.Fields!idbank
        Adodc1.Recordset.Fields!rad = info(12)
        Adodc1.Recordset.Fields!idgarardad = Form2.Adodc1.Recordset.Fields!vamcode
        Adodc1.Recordset.Fields!AllVamAsl = Form2.Adodc1.Recordset.Fields!moneyvamrial
        Adodc1.Recordset.Fields!AllVambahre = Form2.Adodc1.Recordset.Fields!allbahre
        Adodc1.Recordset.Fields!mablaghgestasl = Form9.Adodc2.Recordset.Fields!mablaghgestasl
        Adodc1.Recordset.Fields!mablaghgestbahre = Form9.Adodc2.Recordset.Fields!mablaghgestbahre
        Adodc1.Recordset.Fields!mablaghgestall = Form9.Adodc2.Recordset.Fields!mablaghgestall
        Adodc1.Recordset.Fields!sumpardakhtasl = Form9.Adodc2.Recordset.Fields!sumpardakhtasl
        Adodc1.Recordset.Fields!sumpardakhtbahre = Form9.Adodc2.Recordset.Fields!sumpardakhtbahre
        Adodc1.Recordset.Fields!sumpardakhtall = Form9.Adodc2.Recordset.Fields!sumpardakhtall
        Adodc1.Recordset.Fields!sumbedehiasl = Form9.Adodc2.Recordset.Fields!sumbedehiasl
        Adodc1.Recordset.Fields!sumbedehibahre = Form9.Adodc2.Recordset.Fields!sumbedehibahre
        Adodc1.Recordset.Fields!sumbedehiall = Form9.Adodc2.Recordset.Fields!sumbedehiall
        Adodc1.Recordset.Fields!numberall = Form9.Adodc1.Recordset.RecordCount
        Adodc1.Recordset.Fields!numbermonde = Val(Form9.Adodc1.Recordset.RecordCount) - Val(Form9.Adodc2.Recordset.RecordCount)
        Adodc1.Recordset.Fields!numberdif = Form9.Adodc2.Recordset.RecordCount
        Adodc1.Recordset.Update
      Else
        Adodc1.Refresh
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!rad = info(12)
        Adodc1.Recordset.Fields!idrad = Form2.Adodc1.Recordset.Fields!rad
        Adodc1.Recordset.Fields!idbank = Form2.Adodc1.Recordset.Fields!idbank
        Adodc1.Recordset.Fields!idgarardad = Form2.Adodc1.Recordset.Fields!vamcode
        Adodc1.Recordset.Fields!AllVamAsl = Form2.Adodc1.Recordset.Fields!moneyvamrial
        Adodc1.Recordset.Fields!AllVambahre = Form2.Adodc1.Recordset.Fields!allbahre
        Adodc1.Recordset.Fields!mablaghgestasl = 0
        Adodc1.Recordset.Fields!mablaghgestbahre = 0
        Adodc1.Recordset.Fields!mablaghgestall = 0
        Adodc1.Recordset.Fields!sumpardakhtasl = 0
        Adodc1.Recordset.Fields!sumpardakhtbahre = 0
        Adodc1.Recordset.Fields!sumpardakhtall = 0
        Adodc1.Recordset.Fields!sumbedehiasl = Form2.Adodc1.Recordset.Fields!moneyvamrial
        Adodc1.Recordset.Fields!sumbedehibahre = Form2.Adodc1.Recordset.Fields!allbahre
        Adodc1.Recordset.Fields!sumbedehiall = Val(Form2.Adodc1.Recordset.Fields!moneyvamrial) + Val(Form2.Adodc1.Recordset.Fields!allbahre)
        Adodc1.Recordset.Fields!numberall = Form9.Adodc1.Recordset.RecordCount
        Adodc1.Recordset.Fields!numbermonde = Form9.Adodc1.Recordset.RecordCount
        Adodc1.Recordset.Fields!numberdif = 0
        Adodc1.Recordset.Update
      End If
      info(12) = Val(info(12)) + 1
      Form2.Adodc1.Recordset.MoveNext
    Loop Until Form2.Adodc1.Recordset.EOF = True
  End If
  info(8) = Val(info(2)) - Val(info(5))
  info(9) = Val(info(3)) - Val(info(6))
  info(10) = Val(info(4)) - Val(info(7))
  info(12) = Val(info(12)) + 1
  
  db1.Open Form1.Text1.Text
    rs(0).Open "SELECT Sum(AllVamAsl) as AllVamAsl1 ,Sum(AllVambahre) as AllVambahre1 ,Sum(mablaghgestasl) as mablaghgestasl1 ,Sum(mablaghgestbahre) as mablaghgestbahre1 ,Sum(mablaghgestall) as mablaghgestall1 ,Sum(sumpardakhtasl) as sumpardakhtasl1 ,Sum(sumpardakhtbahre) as sumpardakhtbahre1 ,Sum(sumpardakhtall) as sumpardakhtall1 ,Sum(sumbedehiasl) as sumbedehiasl1 ,Sum(sumbedehibahre) as sumbedehibahre1 ,Sum(sumbedehiall) as sumbedehiall1 FROM printvazriz", db1
        Adodc1.Refresh
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!rad = info(12)
        Adodc1.Recordset.Fields!idgarardad = "ÌãÚ ˜á : "
        Adodc1.Recordset.Fields!AllVamAsl = rs(0).Fields!AllVamAsl1
        Adodc1.Recordset.Fields!AllVambahre = rs(0).Fields!AllVamBahre1
        Adodc1.Recordset.Fields!mablaghgestasl = rs(0).Fields!mablaghgestasl1
        Adodc1.Recordset.Fields!mablaghgestbahre = rs(0).Fields!mablaghgestbahre1
        Adodc1.Recordset.Fields!mablaghgestall = rs(0).Fields!mablaghgestall1
        Adodc1.Recordset.Fields!sumpardakhtasl = rs(0).Fields!sumpardakhtasl1
        Adodc1.Recordset.Fields!sumpardakhtbahre = rs(0).Fields!sumpardakhtbahre1
        Adodc1.Recordset.Fields!sumpardakhtall = rs(0).Fields!sumpardakhtall1
        Adodc1.Recordset.Fields!sumbedehiasl = rs(0).Fields!sumbedehiasl1
        Adodc1.Recordset.Fields!sumbedehibahre = rs(0).Fields!sumbedehibahre1
        Adodc1.Recordset.Fields!sumbedehiall = rs(0).Fields!sumbedehiall1
        Adodc1.Recordset.Fields!numbermonde = ""
        Adodc1.Recordset.Update
    rs(0).Close
  db1.Close
  Adodc1.ConnectionString = Form1.Text1.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM printvazriz ORDER BY rad ASC"
  Adodc1.Refresh
  
  
  Adodc1.Refresh
  DataGrid1.Refresh
  
  Adodc1.Refresh
  DataGrid1.Refresh
  
  Adodc1.Refresh
  DataGrid1.Refresh
  
  Adodc1.Recordset.Sort = "rad"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form7.Show
End Sub

