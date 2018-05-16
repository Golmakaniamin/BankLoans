VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00FFE4CC&
   Caption         =   "»«“ Å—œ«Œ "
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15330
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "À» "
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3840
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ç«Å"
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9960
      Width           =   14895
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   120
      Top             =   1320
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
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE4CC&
      Caption         =   " ‰ŸÌ„«  ﬁ”ÿ"
      Height          =   1575
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   7095
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Text            =   "30"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Text            =   "30"
         Top             =   360
         Width           =   975
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   495
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„œ  “„«‰ Å—œ«Œ  (ò·)"
         Height          =   495
         Index           =   9
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰—Œ «—“"
         Height          =   495
         Index           =   8
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„œ  “„«‰ Å—œ«Œ  („«Â ﬁ»·)"
         Height          =   495
         Index           =   7
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ Å—œ«Œ  ‘„”Ì"
         Height          =   495
         Index           =   6
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   1680
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
      Bindings        =   "Form9.frx":2CFA
      Height          =   5415
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   9551
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "rad2"
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
         DataField       =   "darsad"
         Caption         =   "œ—’œ"
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
         DataField       =   "money"
         Caption         =   "„»·€"
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
         Caption         =   " «—ÌŒ ‘„”Ì"
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
         Caption         =   " «—ÌŒ „Ì·«œÌ"
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
         DataField       =   "datepardakhtall"
         Caption         =   " «—ÌŒ Å—œ«Œ  ‘„”Ì"
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
         DataField       =   "mablaghgestasl"
         Caption         =   "„»·€ ﬁ”ÿ - «’·"
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
         DataField       =   "mablaghgestbahre"
         Caption         =   "„»·€ ﬁ”ÿ - ›—⁄"
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
         DataField       =   "mablaghgestall"
         Caption         =   "„»·€ ﬁ”ÿ - Ã„⁄"
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
         DataField       =   "sumpardakhtasl"
         Caption         =   "„»·€ Å—œ«Œ Ì - «’·"
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
         DataField       =   "sumpardakhtbahre"
         Caption         =   "„»·€ Å—œ«Œ Ì - ›—⁄"
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
         DataField       =   "sumpardakhtall"
         Caption         =   "„»·€ Å—œ«Œ Ì - Ã„⁄"
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
         DataField       =   "sumbedehiasl"
         Caption         =   "„»·€ »œÂÌ - «’·"
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
      BeginProperty Column13 
         DataField       =   "sumbedehibahre"
         Caption         =   "„»·€ »œÂÌ - ›—⁄"
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
      BeginProperty Column14 
         DataField       =   "sumbedehiall"
         Caption         =   "„»·€ »œÂÌ - Ã„⁄"
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
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1484.787
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
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   255
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   5
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4560
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ »«‰ò :"
      Height          =   495
      Index           =   1
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ› Ê«„ :"
      Height          =   495
      Index           =   3
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "Form9.frx":2D0F
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim mApplication  As New Excel.Application
Dim mWorkbook     As New Excel.Workbook
Dim mWorksheet    As New Excel.Worksheet
Dim newq As String

If MaskEdBox1.Text = "" Then
  e = MsgBox("›Ì·œ  «—ÌŒ À»  Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If InStr(MaskEdBox1.Text, "*") <> 0 Then
  e = MsgBox("·ÿ›« ›Ì·œ  «—ÌŒ À»  —«  ò„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If Adodc2.Recordset.RecordCount = Adodc1.Recordset.RecordCount Then
  MsgBox "«ﬁ”«ÿ ‘„« »Â Å«Ì«‰ —”ÌœÂ «” ", vbInformation + vbMsgBoxRight, ""
  Exit Sub
End If

If Adodc1.Recordset.Fields!noee = 0 Then
  Set mWorkbook = mApplication.Workbooks.Add
  Set mWorksheet = mWorkbook.Sheets.Add
  Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount) + 1)), , adSearchForward, 1
  mWorksheet.Name = "BestIrProgramer@yahoo.com"

  With mWorksheet.Cells
     .Cells(1, 1) = (Form6.Adodc1.Recordset.Fields!moneyvamrial / Adodc1.Recordset.RecordCount) * Adodc2.Recordset.RecordCount
     .Cells(2, 1) = Form6.Adodc1.Recordset.Fields!moneyvamrial
     .Cells(3, 1) = Text5.Text
     .Cells(4, 1) = Text3.Text
     .Cells(5, 1) = "=" + Label1(5).Caption
     .Cells(1, 2) = Adodc2.Recordset.RecordCount + 1
     .Cells(2, 2) = Adodc1.Recordset.RecordCount
     .Cells(3, 2) = Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) + Val(Form6.Adodc1.Recordset.Fields!allbahre)
     If Adodc2.Recordset.RecordCount > 0 Then
       Adodc2.Recordset.MoveLast
       .Cells(4, 2) = Adodc2.Recordset.Fields!sumbedehiasl
     Else
       .Cells(4, 2) = Form6.Adodc1.Recordset.Fields!moneyvamrial
     End If
  End With
  
'  mWorkbook.Sheets.Application.Visible = True

  Text2.Text = Int(mWorksheet.Cells(5, 1))
'  MsgBox Int(mWorksheet.Cells(5, 1))
  
  mWorksheet.Delete

  mWorkbook.SaveAs "c:\amingolmakani.09192595284", , , , , , xlNoChange
  mWorkbook.Close
  Kill "c:\amingolmakani.09192595284"
  
  If Adodc2.Recordset.RecordCount = 0 Then
    Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount) + 1)), , adSearchForward, 1

    Text1.Text = 0
    If Form9.DataGrid1.Columns.Item(1).Visible = True Then Text1.Text = (Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) / 100) * Val(Adodc1.Recordset.Fields!darsad)
    If Form9.DataGrid1.Columns.Item(2).Visible = True Then Text1.Text = Adodc1.Recordset.Fields!Money

    Adodc1.Recordset.Fields!mablaghgestall = Text1.Text
    Adodc1.Recordset.Fields!mablaghgestbahre = Text2.Text
    Adodc1.Recordset.Fields!mablaghgestasl = Val(Text1.Text) - Val(Text2.Text)

    Adodc1.Recordset.Fields!sumpardakhtall = Text1.Text
    Adodc1.Recordset.Fields!sumpardakhtbahre = Text2.Text
    Adodc1.Recordset.Fields!sumpardakhtasl = Val(Text1.Text) - Val(Text2.Text)

    Adodc1.Recordset.Fields!sumbedehibahre = Val(Form6.Adodc1.Recordset.Fields!allbahre) - Adodc1.Recordset.Fields!sumpardakhtbahre
    Adodc1.Recordset.Fields!sumbedehiasl = Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) - Adodc1.Recordset.Fields!sumpardakhtasl
    Adodc1.Recordset.Fields!sumbedehiall = Val(Adodc1.Recordset.Fields!sumbedehibahre) + Val(Adodc1.Recordset.Fields!sumbedehiasl)

    Adodc1.Recordset.Fields!datepardakhtall = MaskEdBox1.Text
    Adodc1.Recordset.Fields!YesNo = 1
    Adodc1.Recordset.Update
  Else
    Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount))), , adSearchForward, 1
    q = Adodc1.Recordset.Fields!sumpardakhtasl
    w = Adodc1.Recordset.Fields!sumpardakhtbahre

    Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount) + 1)), , adSearchForward, 1

    Text1.Text = 0
    If Form9.DataGrid1.Columns.Item(1).Visible = True Then Text1.Text = (Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) / 100) * Val(Adodc1.Recordset.Fields!darsad)
    If Form9.DataGrid1.Columns.Item(2).Visible = True Then Text1.Text = Adodc1.Recordset.Fields!Money

    Adodc1.Recordset.Fields!mablaghgestall = Text1.Text
    Adodc1.Recordset.Fields!mablaghgestbahre = Text2.Text
    Adodc1.Recordset.Fields!mablaghgestasl = Val(Text1.Text) - Val(Text2.Text)

    Adodc1.Recordset.Fields!sumpardakhtasl = Val(q) + Val(Adodc1.Recordset.Fields!mablaghgestasl)
    Adodc1.Recordset.Fields!sumpardakhtbahre = Val(w) + Val(Adodc1.Recordset.Fields!mablaghgestbahre)
    Adodc1.Recordset.Fields!sumpardakhtall = Val(Adodc1.Recordset.Fields!sumpardakhtbahre) + Val(Adodc1.Recordset.Fields!sumpardakhtasl)

    Adodc1.Recordset.Fields!sumbedehiall = (Val(Form6.Adodc1.Recordset.Fields!allbahre) + Val(Form6.Adodc1.Recordset.Fields!moneyvamrial)) - Val(Adodc1.Recordset.Fields!sumpardakhtall)
    Adodc1.Recordset.Fields!sumbedehibahre = Val(Form6.Adodc1.Recordset.Fields!allbahre) - Adodc1.Recordset.Fields!sumpardakhtbahre
    Adodc1.Recordset.Fields!sumbedehiasl = (Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) - Adodc1.Recordset.Fields!sumpardakhtasl)

    Adodc1.Recordset.Fields!datepardakhtall = MaskEdBox1.Text
    Adodc1.Recordset.Fields!YesNo = 1
    'tasri
    Adodc1.Recordset.Update
    Adodc1.Refresh
  End If
  'noee=1
Else
  Set mWorkbook = mApplication.Workbooks.Add
  Set mWorksheet = mWorkbook.Sheets.Add
  Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount) + 1)), , adSearchForward, 1
  mWorksheet.Name = "BestIrProgramer@yahoo.com"

  If Adodc2.Recordset.EOF = True Then
    newq = Form6.Adodc1.Recordset.Fields!moneyvamrial
  Else
    newq = Adodc2.Recordset.Fields!sumbedehiasl
  End If
  With mWorksheet.Cells
     .Cells(1, 1) = newq
     .Cells(2, 1) = Form6.Adodc1.Recordset.Fields!moneyvamrial
     .Cells(3, 1) = Text5.Text
     .Cells(4, 1) = Text3.Text
     .Cells(5, 1) = "=" + Label1(5).Caption
     .Cells(1, 2) = Adodc2.Recordset.RecordCount + 1
     .Cells(2, 2) = Adodc1.Recordset.RecordCount
     .Cells(3, 2) = Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) + Val(Form6.Adodc1.Recordset.Fields!allbahre)
     If Adodc2.Recordset.RecordCount > 0 Then
       Adodc2.Recordset.MoveLast
       .Cells(4, 2) = Adodc2.Recordset.Fields!sumbedehiasl
     Else
       .Cells(4, 2) = Form6.Adodc1.Recordset.Fields!moneyvamrial
     End If
     
  End With

  'mWorkbook.Sheets.Application.Visible = True

  Text2.Text = Int(mWorksheet.Cells(5, 1))
  '  MsgBox Int(mWorksheet.Cells(5, 1))
  mWorksheet.Delete

  mWorkbook.SaveAs "c:\amingolmakani.09192595284", , , , , , xlNoChange
  mWorkbook.Close
  Kill "c:\amingolmakani.09192595284"
  
  If Adodc2.Recordset.RecordCount = 0 Then
    Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount) + 1)), , adSearchForward, 1

    Text1.Text = 0
    If Form9.DataGrid1.Columns.Item(1).Visible = True Then Text1.Text = (Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) / 100) * Val(Adodc1.Recordset.Fields!darsad)
    If Form9.DataGrid1.Columns.Item(2).Visible = True Then Text1.Text = Adodc1.Recordset.Fields!Money

    Adodc1.Recordset.Fields!mablaghgestall = Val(Text1.Text) + Val(Text2.Text)
    Adodc1.Recordset.Fields!mablaghgestbahre = Text2.Text
    Adodc1.Recordset.Fields!mablaghgestasl = Text1.Text

    Adodc1.Recordset.Fields!sumpardakhtall = Val(Text1.Text) + Val(Text2.Text)
    Adodc1.Recordset.Fields!sumpardakhtbahre = Text2.Text
    Adodc1.Recordset.Fields!sumpardakhtasl = Text1.Text

    Adodc1.Recordset.Fields!sumbedehibahre = Val(Form6.Adodc1.Recordset.Fields!allbahre) - Adodc1.Recordset.Fields!sumpardakhtbahre
    Adodc1.Recordset.Fields!sumbedehiasl = Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) - Adodc1.Recordset.Fields!sumpardakhtasl
    Adodc1.Recordset.Fields!sumbedehiall = Val(Adodc1.Recordset.Fields!sumbedehibahre) + Val(Adodc1.Recordset.Fields!sumbedehiasl)

    Adodc1.Recordset.Fields!datepardakhtall = MaskEdBox1.Text
    Adodc1.Recordset.Fields!YesNo = 1
    Adodc1.Recordset.Update
  Else
    Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount))), , adSearchForward, 1
    q = Adodc1.Recordset.Fields!sumpardakhtasl
    w = Adodc1.Recordset.Fields!sumpardakhtbahre

    Adodc1.Recordset.Find "rad2=" + Trim(Str(Val(Adodc2.Recordset.RecordCount) + 1)), , adSearchForward, 1

    Text1.Text = 0
    If Form9.DataGrid1.Columns.Item(1).Visible = True Then Text1.Text = (Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) / 100) * Val(Adodc1.Recordset.Fields!darsad)
    If Form9.DataGrid1.Columns.Item(2).Visible = True Then Text1.Text = Adodc1.Recordset.Fields!Money

    Adodc1.Recordset.Fields!mablaghgestall = Val(Text1.Text) + Val(Text2.Text)
    Adodc1.Recordset.Fields!mablaghgestbahre = Text2.Text
    Adodc1.Recordset.Fields!mablaghgestasl = Text1.Text

    Adodc1.Recordset.Fields!sumpardakhtasl = Val(q) + Val(Adodc1.Recordset.Fields!mablaghgestasl)
    Adodc1.Recordset.Fields!sumpardakhtbahre = Val(w) + Val(Adodc1.Recordset.Fields!mablaghgestbahre)
    Adodc1.Recordset.Fields!sumpardakhtall = Val(Adodc1.Recordset.Fields!sumpardakhtbahre) + Val(Adodc1.Recordset.Fields!sumpardakhtasl)

    Adodc1.Recordset.Fields!sumbedehiall = (Val(Form6.Adodc1.Recordset.Fields!allbahre) + Val(Form6.Adodc1.Recordset.Fields!moneyvamrial)) - Val(Adodc1.Recordset.Fields!sumpardakhtall)
    Adodc1.Recordset.Fields!sumbedehibahre = Val(Form6.Adodc1.Recordset.Fields!allbahre) - Adodc1.Recordset.Fields!sumpardakhtbahre
    Adodc1.Recordset.Fields!sumbedehiasl = (Val(Form6.Adodc1.Recordset.Fields!moneyvamrial) - Adodc1.Recordset.Fields!sumpardakhtasl)

    Adodc1.Recordset.Fields!datepardakhtall = MaskEdBox1.Text
    Adodc1.Recordset.Fields!YesNo = 1
    Adodc1.Recordset.Update
    Adodc1.Refresh
  End If
End If

DataGrid1.Refresh
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Form6.Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Form6.Adodc1.Recordset.Fields!rad)) + ") ORDER BY rad2"
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.Refresh
DataGrid1.Refresh

Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Form6.Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Form6.Adodc1.Recordset.Fields!rad)) + ") and (yesno='1') ORDER BY rad2"
Adodc2.Refresh
Adodc2.Refresh
DataGrid1.Refresh

If Adodc2.Recordset.RecordCount = Adodc1.Recordset.RecordCount Then
  MsgBox "«ﬁ”«ÿ ‘„« »Â Å«Ì«‰ —”ÌœÂ «” ", vbInformation + vbMsgBoxRight, ""
  Form6.Adodc1.Recordset.Fields!tasvie = "‘œÂ"
  Form6.Adodc1.Recordset.Update
End If
MaskEdBox1.SetFocus
End Sub


Private Sub Command2_Click()
Form14.Show
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  e = MsgBox("¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
  If e = 6 Then
    Adodc1.Recordset.Fields!mablaghgestall = ""
    Adodc1.Recordset.Fields!mablaghgestbahre = ""
    Adodc1.Recordset.Fields!mablaghgestasl = ""

    Adodc1.Recordset.Fields!sumpardakhtall = ""
    Adodc1.Recordset.Fields!sumpardakhtbahre = ""
    Adodc1.Recordset.Fields!sumpardakhtasl = ""

    Adodc1.Recordset.Fields!sumbedehibahre = ""
    Adodc1.Recordset.Fields!sumbedehiasl = ""
    Adodc1.Recordset.Fields!sumbedehiall = ""

    Adodc1.Recordset.Fields!datepardakhtall = ""
    Adodc1.Recordset.Fields!YesNo = 0
    Adodc1.Recordset.Update
    
    Adodc2.CommandType = adCmdUnknown
    Adodc2.RecordSource = "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Form6.Adodc1.Recordset.Fields!idbank)) + ") and (rad=" + Trim(Str(Form6.Adodc1.Recordset.Fields!rad)) + ") and (yesno='1') ORDER BY rad2"
    Adodc2.Refresh
    DataGrid1.Refresh
  End If
End If
End Sub

Private Sub Form_Activate()
MaskEdBox1.Mask = "13##/##/##"
MaskEdBox1.PromptChar = "*"
MaskEdBox1.Text = Amin.mil2shams(Date$)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Label2.Caption = 1 Then
  Form6.Show
  Me.Hide
Else
  Form8.Show
  Me.Hide
End If
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub

Private Sub MaskEdBox1_LostFocus()
If Adodc2.Recordset.RecordCount = 0 Then
  Text3.Text = DateDiff("d", Form6.Adodc1.Recordset.Fields!datemiladi, Amin.shams2mil(MaskEdBox1.Text))
  Text5.Text = DateDiff("d", Form6.Adodc1.Recordset.Fields!datemiladi, Amin.shams2mil(MaskEdBox1.Text))
Else
  Adodc2.Recordset.MoveLast
  Text3.Text = DateDiff("d", Adodc2.Recordset.Fields!datemiladi, Amin.shams2mil(MaskEdBox1.Text))
  Text5.Text = DateDiff("d", Form6.Adodc1.Recordset.Fields!datemiladi, Amin.shams2mil(MaskEdBox1.Text))
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub
