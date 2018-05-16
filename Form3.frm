VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFE4CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ÕÊÂ »«“Å—œ«Œ  Ê«„"
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":2CFA
      Height          =   8175
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   14420
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   6
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
         Caption         =   " «—ÌŒ û „Ì·«œÌ"
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
         DataField       =   "money1"
         Caption         =   "„»·€ ò«—„“œ »«‰ò"
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
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE4CC&
      Caption         =   "‰ÕÊÂ Å—œ«Œ "
      Height          =   1095
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   4215
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFE4CC&
         Caption         =   "œ—’œÌ"
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFE4CC&
         Caption         =   "—Ì«·Ì"
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFE4CC&
      Caption         =   "„‘Œ’«  Å—œ«Œ  Â«"
      Height          =   6015
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4320
      Width           =   4215
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFE4CC&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   3360
         Width           =   2535
         Begin VB.OptionButton Option6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFE4CC&
            Caption         =   "¬Œ—Ì‰"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFE4CC&
            Caption         =   "«Ê·Ì‰ ﬁ”ÿ"
            Height          =   375
            Left            =   1320
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "À» "
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   5400
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Text            =   "1"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFE4CC&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2640
         Width           =   2535
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFE4CC&
            Caption         =   "«’· ﬁ”ÿ"
            Height          =   375
            Left            =   1320
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFE4CC&
            Caption         =   "ò· ﬁ”ÿ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «ÀÌ— ﬁ”ÿ ‰«„”«ÊÌ :"
         Height          =   495
         Index           =   12
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «ÀÌ— „»·€ ﬁ”ÿ œ— :"
         Height          =   495
         Index           =   11
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ „«ÂÂ«Ì „Ì«‰ »«“ Å—œ«Œ "
         Height          =   495
         Index           =   10
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ﬁ”«ÿ ‰« „”«ÊÌ"
         Height          =   495
         Index           =   9
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ «ﬁ”«ÿ „”«ÊÌ"
         Height          =   495
         Index           =   8
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ «ﬁ”«ÿ"
         Height          =   495
         Index           =   7
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ‘—Ê⁄ »«“Å—œ«Œ  ‘„”Ì"
         Height          =   495
         Index           =   6
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ‘—Ê⁄ »«“Å—œ«Œ  „Ì·«œÌ"
         Height          =   495
         Index           =   5
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ› :"
      Height          =   495
      Index           =   3
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ »«‰ò :"
      Height          =   495
      Index           =   1
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   8520
      Left            =   0
      Picture         =   "Form3.frx":2D0F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11145
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoConnection1 As ADODB.Connection
Dim cmd1 As ADODB.Command
Dim adoRecordset1 As ADODB.Recordset

Private Sub Command1_Click()
Dim mApplication  As New Excel.Application
Dim mWorkbook     As New Excel.Workbook
Dim mWorksheet    As New Excel.Worksheet
Dim newq As String, newq1 As String

If MaskEdBox1.Text = "" Then
  e = MsgBox("›Ì·œ  «—ÌŒ ‘—Ê⁄ »«“Å—œ«Œ  ‘„”Ì Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If InStr(MaskEdBox1.Text, "*") <> 0 Then
  e = MsgBox("›Ì·œ  «—ÌŒ ‘—Ê⁄ »«“Å—œ«Œ  ‘„”Ì —«  ﬂ„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox1.SetFocus
  Exit Sub
End If

If MaskEdBox2.Text = "" Then
  e = MsgBox("›Ì·œ  «—ÌŒ ‘—Ê⁄ »«“Å—œ«Œ  „Ì·«œÌ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox2.SetFocus
  Exit Sub
End If

If InStr(MaskEdBox2.Text, "*") <> 0 Then
  e = MsgBox("›Ì·œ  «—ÌŒ ‘—Ê⁄ »«“Å—œ«Œ  „Ì·«œÌ —«  ﬂ„Ì· ‰„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  MaskEdBox2.SetFocus
  Exit Sub
End If

If Text1.Text = "" Then
  e = MsgBox("›Ì·œ  ⁄œ«œ «ﬁ”«ÿ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1.SetFocus
  Exit Sub
End If

If Text4.Text = "" Then
  e = MsgBox("›Ì·œ  ⁄œ«œ „«ÂÂ«Ì „Ì«‰ »«“ Å—œ«Œ  Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text4.SetFocus
  Exit Sub
End If

If Text4.Text = 0 Then
  e = MsgBox("›Ì·œ  ⁄œ«œ „«ÂÂ«Ì „Ì«‰ »«“ Å—œ«Œ  ‰„Ì  Ê«‰œ ’›— »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text4.SetFocus
  Exit Sub
End If

If Text2.Text = "" Then
  e = MsgBox("›Ì·œ „»·€ «ﬁ”«ÿ „”«ÊÌ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1.SetFocus
  Exit Sub
End If

If Text3.Text = "" Then
  e = MsgBox("›Ì·œ „»·€ «ﬁ”«ÿ ‰« „”«ÊÌ Œ«·Ì „Ì »«‘œ" & Chr(13) & Chr(10) & "·ÿ›« ‰”»  »Â  ò„Ì· ò—œ‰ ¬‰ «ﬁœ«„ ›—„«ÌÌœ", vbCritical + vbMsgBoxRight, "")
  Text1.SetFocus
  Exit Sub
End If

e = MsgBox("œ— ’Ê—   «ÌÌœ «ÿ·«⁄«  ﬁ»·Ì Å«ﬂ „Ì ‘Êœ" + Chr(10) + Chr(13) + "¬Ì« ‘„« „ÿ„∆‰ Â” Ìœ", vbCritical + vbMsgBoxRight + vbYesNo, "")
If e = 6 Then
  If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveFirst
    Do
      Adodc1.Recordset.Delete
      Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
  End If
  ww = 0
  For q = 1 To Val(Text1.Text) * Val(Text4.Text) Step Val(Text4.Text)
    ww = Val(ww) + Val(1)
    Adodc1.Refresh
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!idbank = Label1(4).Caption
    Adodc1.Recordset.Fields!rad = Label1(2).Caption
    Adodc1.Recordset.Fields!rad2 = ww
    If Option1.Value = True Then
      Adodc1.Recordset.Fields!darsad = 0
      If Option6.Value = True Then
        If q <> Val(Text1.Text) Then
          Adodc1.Recordset.Fields!Money = Text2.Text
        Else
          Adodc1.Recordset.Fields!Money = Text3.Text
        End If
      End If
      If Option5.Value = True Then
        If q = 1 Then
          Adodc1.Recordset.Fields!Money = Text3.Text
        Else
          Adodc1.Recordset.Fields!Money = Text2.Text
        End If
      End If
    Else
      Adodc1.Recordset.Fields!darsad = 0
      Adodc1.Recordset.Fields!Money = 0
    End If
    Adodc1.Recordset.Fields!money1 = 0
    If Option3.Value = True Then Adodc1.Recordset.Fields!noee = 1
    If Option4.Value = True Then Adodc1.Recordset.Fields!noee = 0
    Adodc1.Recordset.Fields!dateshamsi = Amin.dateaminEzafeMoon(MaskEdBox1.Text, Trim(Str(q - 1)))
    Adodc1.Recordset.Fields!datemiladi = Amin.shams2mil(Adodc1.Recordset.Fields!dateshamsi)
    Adodc1.Recordset.Fields!Groupme = Form1.Combo2.List(Form1.Combo1.ListIndex)
    Adodc1.Recordset.Update
    Adodc1.Refresh
  Next q
  
  Adodc1.Refresh
  DataGrid1.Refresh
  
  Adodc1.Refresh
  DataGrid1.Refresh
  
  Adodc1.Recordset.Sort = "rad2"
End If
End Sub

Private Sub Form_Activate()
MaskEdBox1.Mask = "13##/##/##"
MaskEdBox1.PromptChar = "*"

MaskEdBox2.Mask = "20##/##/##"
MaskEdBox2.PromptChar = "*"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Show
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
If KeyCode = 13 Then Text1.SetFocus
End Sub

Private Sub MaskEdBox2_LostFocus()
If InStr(MaskEdBox2.Text, "*") = 0 Then
  Dim aa1 As String
  aa1 = Mid(MaskEdBox2.Text, 6, 2) + "-" + Mid(MaskEdBox2.Text, 9, 2) + "-" + Mid(MaskEdBox2.Text, 1, 4)
  MaskEdBox1.Text = Amin.mil2shams(aa1)
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Option1_Click()
DataGrid1.Columns.Item(1).Visible = False
DataGrid1.Columns.Item(2).Visible = True
Command1.Top = 5400
Frame2.Height = 6015
DataGrid1.Height = 8175
Form3.Height = 10890
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Option2_Click()
DataGrid1.Columns.Item(1).Visible = True
DataGrid1.Columns.Item(2).Visible = False
Command1.Top = 4200
Frame2.Height = 4815
DataGrid1.Height = 6975
Form3.Height = 9720
Text2.Text = "-"
Text3.Text = "-"
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If Option1.Value = True Then If KeyCode = 13 Then Text2.SetFocus
If Option2.Value = True Then If KeyCode = 13 Then Command1.SetFocus
End Sub
