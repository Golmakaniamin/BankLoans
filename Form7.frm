VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00FFE4CC&
   Caption         =   "ÒÇÑÔÇÊ"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "ÂÎÑíä æÖÚíÊ ÊÓåíáÇÊ ÈÇä˜åÇ"
      TabPicture(0)   =   "Form7.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command7"
      Tab(0).Control(1)=   "Command3"
      Tab(0).Control(2)=   "DataGrid1"
      Tab(0).Control(3)=   "Adodc1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "ÊÚåÏÇÊ Èå ÈÇä˜ "
      TabPicture(1)   =   "Form7.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Adodc2"
      Tab(1).Control(1)=   "Command6"
      Tab(1).Control(2)=   "Combo1"
      Tab(1).Control(3)=   "Combo2"
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(5)=   "MaskEdBox1"
      Tab(1).Control(6)=   "MaskEdBox2"
      Tab(1).Control(7)=   "DataGrid2"
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(9)=   "Label1(3)"
      Tab(1).Control(10)=   "Label1(2)"
      Tab(1).Control(11)=   "Label1(0)"
      Tab(1).Control(12)=   "Label1(1)"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "äãæÏÇÑ ÊÚåÏÇÊ ãÇåÇäå"
      TabPicture(2)   =   "Form7.frx":2D32
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(4)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(5)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(7)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(8)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Combo3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Combo4"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Command2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "MaskEdBox3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Picture1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "ÊÚåÏÇÊ ÈÇä˜ åÇ ãÇåÇäå"
      TabPicture(3)   =   "Form7.frx":2D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command5"
      Tab(3).Control(1)=   "Option2"
      Tab(3).Control(2)=   "Option1"
      Tab(3).Control(3)=   "Command4"
      Tab(3).Control(4)=   "MaskEdBox4"
      Tab(3).Control(5)=   "DataGrid3"
      Tab(3).Control(6)=   "Adodc3"
      Tab(3).Control(7)=   "Label1(6)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "ãÌãæÚ æÇã åÇ ÑíÇáí"
      TabPicture(4)   =   "Form7.frx":2D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label3"
      Tab(4).Control(1)=   "Label1(9)"
      Tab(4).Control(2)=   "Label1(10)"
      Tab(4).Control(3)=   "Label1(11)"
      Tab(4).Control(4)=   "Label1(12)"
      Tab(4).Control(5)=   "DataGrid4"
      Tab(4).Control(6)=   "MaskEdBox6"
      Tab(4).Control(7)=   "MaskEdBox5"
      Tab(4).Control(8)=   "Command8"
      Tab(4).Control(9)=   "Combo6"
      Tab(4).Control(10)=   "Combo5"
      Tab(4).Control(11)=   "Command9"
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "ãÇäÏå ÇÈÊÏÇí ÓÇá"
      TabPicture(5)   =   "Form7.frx":2D86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(13)"
      Tab(5).Control(1)=   "Label1(14)"
      Tab(5).Control(2)=   "DataGrid5"
      Tab(5).Control(3)=   "Combo8"
      Tab(5).Control(4)=   "MaskEdBox7"
      Tab(5).Control(5)=   "Command10"
      Tab(5).Control(6)=   "Command11"
      Tab(5).Control(7)=   "Combo7"
      Tab(5).ControlCount=   8
      Begin VB.ComboBox Combo7 
         Height          =   465
         Left            =   -64200
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Text            =   "Combo2"
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Ç"
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   600
         Width           =   3855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "äãÇíÔ"
         Height          =   495
         Left            =   -70680
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   600
         Width           =   3855
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -74880
         Top             =   1500
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
      Begin VB.CommandButton Command9 
         Caption         =   "äãÇíÔ"
         Height          =   495
         Left            =   -71760
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   540
         Width           =   2895
      End
      Begin VB.ComboBox Combo5 
         Height          =   465
         Left            =   -64200
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Text            =   "Combo2"
         Top             =   540
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo6 
         Height          =   465
         ItemData        =   "Form7.frx":2DA2
         Left            =   -64200
         List            =   "Form7.frx":2DC7
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   540
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Ç"
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   540
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "äãÇíÔ"
         Height          =   495
         Left            =   -68160
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   6495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ç"
         Height          =   495
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   540
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ç"
         Height          =   495
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   4095
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "ÐÔÊå"
         Height          =   495
         Left            =   -66120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÂíäÏå"
         Height          =   495
         Left            =   -65040
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "äãÇíÔ"
         Height          =   495
         Left            =   -70680
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   4095
      End
      Begin VB.PictureBox Picture1 
         Height          =   7095
         Left            =   120
         RightToLeft     =   -1  'True
         ScaleHeight     =   7035
         ScaleWidth      =   13155
         TabIndex        =   20
         Top             =   1320
         Width           =   13215
         Begin VB.HScrollBar HScroll1 
            Height          =   375
            Left            =   0
            Max             =   0
            Min             =   -15960
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   6600
            Width           =   13095
         End
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   6855
            Left            =   0
            OleObjectBlob   =   "Form7.frx":2E2F
            TabIndex        =   22
            Top             =   0
            Width           =   29160
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ç"
         Height          =   495
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   6495
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   495
         Left            =   8400
         TabIndex        =   17
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command2 
         Caption         =   "äãÇíÔ"
         Height          =   495
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   5895
      End
      Begin VB.ComboBox Combo4 
         Height          =   465
         Left            =   10920
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "Combo2"
         Top             =   540
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         Height          =   465
         ItemData        =   "Form7.frx":50A5
         Left            =   10920
         List            =   "Form7.frx":50CA
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   540
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   465
         ItemData        =   "Form7.frx":5132
         Left            =   -64200
         List            =   "Form7.frx":5157
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   465
         Left            =   -64080
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   780
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "äãÇíÔ"
         Height          =   495
         Left            =   -71760
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form7.frx":51BF
         Height          =   7335
         Left            =   -74880
         TabIndex        =   1
         Top             =   1080
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   25
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
            DataField       =   "id"
            Caption         =   "ÑÏíÝ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "name"
            Caption         =   "äÇã ÈÇä˜"
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
            DataField       =   "Asl_VAM_daryaftani"
            Caption         =   "ÇÕá æÇã åÇí ÏÑíÇÝÊí"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Bahre_VAM_daryaftani"
            Caption         =   "˜ÇÑãÒÏ æÇã åÇí ÏÑíÇÝÊí"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "SUM_VAM_daryaftani"
            Caption         =   "ÌãÚ ÇÕá æ ÈåÑå æÇã åÇí ÏÑíÇÝÊí"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Asl_aghsat_pardakhti"
            Caption         =   "ÇÕá ÇÞÓÇØ ÑÏÇÎÊ ÔÏå"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Bahre_aghsat_pardakhti"
            Caption         =   "˜ÇÑãÒÏ ÇÞÓÇØ ÑÏÇÎÊ ÔÏå"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "SUM_aghsat_pardakhti"
            Caption         =   "ÌãÚ ÇÕá æ ÈåÑå ÇÞÓÇØ ÑÏÇÎÊ ÔÏå"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Asl_Monde_aghsat_daryaftani"
            Caption         =   "ãÇäÏå ÇÕá æÇã åÇí ÊÓæíå äÔÏå"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Bahre_Monde_aghsat_daryaftani"
            Caption         =   "ãÇäÏå ˜ÇÑãÒÏ æÇã åÇí ÊÓæíå äÔÏå"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "SUM_Monde_aghsat_daryaftani"
            Caption         =   "ãÇäÏå ÇÕá æ ÈåÑå æÇã åÇí ÊÓæíå äÔÏå"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "saghf"
            Caption         =   "˜á ÓÞÝ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "sagfe_baz"
            Caption         =   "ÓÞÝ ÈÇÒ ÈÇä˜"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1679.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1785.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1890.142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1980.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2174.74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1874.835
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2190.047
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2129.953
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   2385.071
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   2039.811
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   -74880
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   495
         Left            =   -66720
         TabIndex        =   4
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   495
         Left            =   -68760
         TabIndex        =   5
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form7.frx":51D4
         Height          =   6615
         Left            =   -74880
         TabIndex        =   7
         Top             =   1320
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   11668
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   23
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
         ColumnCount     =   4
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
            DataField       =   "idgharardad"
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
            DataField       =   "money"
            Caption         =   "ãÈáÛ"
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
         BeginProperty Column03 
            DataField       =   "enddatepardakht"
            Caption         =   "ÂÎÑíä ÊÇÑíÎ ÑÏÇÎÊ"
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1725.165
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   495
         Left            =   -63960
         TabIndex        =   23
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form7.frx":51E9
         Height          =   7335
         Left            =   -74880
         TabIndex        =   28
         Top             =   1080
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   25
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
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "ÑÏíÝ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "bankname"
            Caption         =   "äÇã ÈÇä˜"
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
            DataField       =   "m1"
            Caption         =   "ÝÑæÑÏíä"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "m2"
            Caption         =   "ÇÑÏíÈåÔÊ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "m3"
            Caption         =   "ÎÑÏÇÏ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "m4"
            Caption         =   "ÊíÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "m5"
            Caption         =   "ãÑÏÇÏ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "m6"
            Caption         =   "ÔåÑíæÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "m7"
            Caption         =   "ãåÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "m8"
            Caption         =   "ÂÈÇä"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "m9"
            Caption         =   "ÂÐÑ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "m10"
            Caption         =   "Ïí"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "m11"
            Caption         =   "Èåãä"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1080
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "m12"
            Caption         =   "ÇÓÝäÏ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """ÑíÇá"" #,##0"
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
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   -74880
         Top             =   8040
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   495
         Left            =   -66720
         TabIndex        =   38
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   495
         Left            =   -68760
         TabIndex        =   39
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Form7.frx":51FE
         Height          =   6615
         Left            =   -74760
         TabIndex        =   45
         Top             =   1320
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   11668
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
            Caption         =   "ÑÏíÝ"
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
            Caption         =   "ßÏ æÇã"
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
            Caption         =   "äæÚ æÇã"
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
            Caption         =   "ÊÇÑíÎ ÏÑíÇÝÊ ÔãÓí"
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
            Caption         =   "ÊÇÑíÎ ÏÑíÇÝÊ ãíáÇÏí"
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
            Caption         =   "ãÈáÛ æÇã ÇÑÒí"
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
            Caption         =   "äÑÎ ÇÑÒ"
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
            Caption         =   "äæÚ ÇÑÒ"
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
            Caption         =   "ãÈáÛ æÇã ÑíÇáí"
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
            Caption         =   "ãÈáÛ ÈåÑå ßá æÇã ÑíÇáí"
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
            Caption         =   "ÔãÇÑå ÝÑãæá"
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
            Caption         =   "ãæÖæÚ"
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
            Caption         =   "ÊÓæíå"
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   495
         Left            =   -66720
         TabIndex        =   46
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo8 
         Height          =   465
         ItemData        =   "Form7.frx":5213
         Left            =   -64200
         List            =   "Form7.frx":5238
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   600
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "Form7.frx":52A0
         Height          =   6615
         Left            =   -74760
         TabIndex        =   53
         Top             =   1200
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   11668
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
         ColumnCount     =   5
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
            DataField       =   "bankname"
            Caption         =   "˜Ï æÇã"
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
            DataField       =   "darvamasl"
            Caption         =   "ãÈáÛ æÇã"
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
            DataField       =   "darvambahre"
            Caption         =   "ÑÏÇÎÊí"
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
            DataField       =   "darvamsum"
            Caption         =   "ÈÏåí"
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2025.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1830.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2415.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2580.095
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "äÇã ÈÇä˜"
         Height          =   495
         Index           =   14
         Left            =   -62520
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÓÇá"
         Height          =   495
         Index           =   13
         Left            =   -65040
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÒ ÊÇÑíÎ"
         Height          =   495
         Index           =   12
         Left            =   -65040
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÇ"
         Height          =   495
         Index           =   11
         Left            =   -67080
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "äÇã ÈÇä˜"
         Height          =   495
         Index           =   10
         Left            =   -62520
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ãÌãæÚ :"
         Height          =   495
         Index           =   9
         Left            =   -62400
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   7980
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   -67800
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   7980
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   495
         Index           =   8
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÚÏÇÏ æÇã åÇ :"
         Height          =   495
         Index           =   7
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÇÑíÎ"
         Height          =   495
         Index           =   6
         Left            =   -62400
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÇÑíÎ"
         Height          =   495
         Index           =   5
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "äÇã ÈÇä˜"
         Height          =   495
         Index           =   4
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   -67800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   7980
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ãÌãæÚ :"
         Height          =   495
         Index           =   3
         Left            =   -62400
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   7980
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "äÇã ÈÇä˜"
         Height          =   495
         Index           =   2
         Left            =   -62520
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÇ"
         Height          =   495
         Index           =   0
         Left            =   -67080
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÒ ÊÇÑíÎ"
         Height          =   495
         Index           =   1
         Left            =   -65040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   360
      Top             =   1200
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   0
      Top             =   0
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
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Account_vam"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Account_vam"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "printvaz"
      Caption         =   "Adodc4"
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
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim info(13) As String

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox1.SetFocus
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox3.SetFocus
End Sub

Private Sub Combo6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox5.SetFocus
End Sub

Private Sub Command1_Click()
Dim a1 As String, a2 As String, a3 As String, a4 As String
Dim S As String, w As Integer, e As String, r As String
Dim numberall1 As Integer, numberall2 As Integer, numberall3 As Integer
db1.Open Form1.Text1.Text
rs1.Open "DELETE FROM printtaahod", db1
db1.Close

db1.Open Form1.Text1.Text
rs1.Open "SELECT count(idbank) as count1 FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Combo2.List(Combo1.ListIndex) + ")and(tasvie='äÔÏå')", db1
  numberall1 = rs1.Fields!count1
rs1.Close

If numberall1 > 0 Then
  rs1.Open "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Combo2.List(Combo1.ListIndex) + ")and(tasvie='äÔÏå')", db1
    w = 0
    rs1.MoveFirst
    Do
      DoEvents
      a1 = MaskEdBox1.Text
      a2 = MaskEdBox2.Text
      a3 = rs1.Fields!dateshamsi
      
      rs2.Open "SELECT count(idbank) as count1 FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!idbank)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") and (yesno='1')", db1
        numberall2 = rs2.Fields!count1
      rs2.Close
    
      If numberall2 > 0 Then
        rs3.Open "SELECT COUNT(Money) as countMoney, SUM(Money) as Money11,SUM(Money1) as Money12 FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!idbank)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") AND (dateshamsi>='" + MaskEdBox1.Text + "') and (dateshamsi<='" + MaskEdBox2.Text + "') and (yesno='0')", db1
        If rs3.Fields!countMoney > 0 Then
          rs2.Open "SELECT Max(rad2) as maxrad FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!idbank)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") and (yesno='1')", db1
          rs4.Open "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!idbank)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") and (yesno='1') and (rad2=" + Trim(Str(rs2.Fields!maxrad)) + ")", db1
          e = (Val(rs3.Fields!Money11) + Val(rs3.Fields!money12))
          If e > 0 Then
            w = w + 1
            Adodc2.Refresh
            Adodc2.Recordset.AddNew
            Adodc2.Recordset.Fields!rad = w
            Adodc2.Recordset.Fields!idgharardad = rs1.Fields!vamcode
            Adodc2.Recordset.Fields!Money = e
            Adodc2.Recordset.Fields!enddatepardakht = rs4.Fields!dateshamsi
            Adodc2.Recordset.Update
          End If
          rs2.Close
          rs4.Close
        End If
        rs3.Close
      Else
        rs3.Open "SELECT COUNT(Money) as countMoney, SUM(Money) as Money11,SUM(Money1) as Money12 FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!idbank)) + ") and (rad=" + Trim(Str(rs1.Fields!rad)) + ") AND (dateshamsi>='" + MaskEdBox1.Text + "') and (dateshamsi<='" + MaskEdBox2.Text + "') and (yesno='0')", db1
          If rs3.Fields!countMoney > 0 Then
            e = (Val(rs3.Fields!Money11) + Val(rs3.Fields!money12))
            w = w + 1
            Adodc2.Refresh
            Adodc2.Recordset.AddNew
            Adodc2.Recordset.Fields!rad = w
            Adodc2.Recordset.Fields!idgharardad = rs1.Fields!vamcode
            Adodc2.Recordset.Fields!Money = e
            Adodc2.Recordset.Fields!enddatepardakht = "ÑÏÇÎÊ Çæá"
            Adodc2.Recordset.Update
          End If
        rs3.Close
      End If
    rs1.MoveNext
  Loop Until rs1.EOF = True
  
  Adodc2.Refresh
  Adodc2.Refresh
  Adodc2.Recordset.Sort = "rad"
  Label2.Caption = 0
  If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveFirst
    Do
      Label2.Caption = Val(Label2.Caption) + Val(Adodc2.Recordset.Fields!Money)
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
    Label2.Caption = Amin.moneyaminjoda(Label2.Caption)
  End If
  
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "SELECT * FROM printtaahod"
  Adodc2.Refresh
  
  DataGrid2.Refresh
  Adodc2.Refresh
  
  DataGrid2.Refresh
  Adodc2.Recordset.Sort = "rad"
  rs1.Close
End If
db1.Close
End Sub

Private Sub Command10_Click()
If Combo8.ListIndex <> -1 Then
  db1.Open Form1.Text1.Text
    rs1.Open "DELETE FROM printvaz", db1
  db1.Close
  Adodc5.ConnectionString = Form1.Text1.Text
  Adodc5.CommandType = adCmdUnknown
  Adodc5.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo7.List(Combo8.ListIndex))) + ") And (Left(dateshamsi,4)='" + MaskEdBox7.Text + "')"
  Adodc5.Refresh
  If Adodc5.Recordset.RecordCount > 0 Then
    Adodc5.ConnectionString = Form1.Text1.Text
    Adodc5.CommandType = adCmdUnknown
    Adodc5.RecordSource = "SELECT * FROM printvaz"
    Adodc5.Refresh
    q = 1
    db1.Open Form1.Text1.Text
      rs1.Open "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo7.List(Combo8.ListIndex))) + ") And (Left(dateshamsi,4)='" + MaskEdBox7.Text + "')", db1
        rs1.MoveFirst
        Do
          Adodc5.Recordset.AddNew
          Adodc5.Recordset.Fields!rad = q
          Adodc5.Recordset.Fields!bankname = rs1.Fields!vamcode
          Adodc5.Recordset.Fields!darvamasl = rs1.Fields!moneyvamrial + rs1.Fields!allbahre
          
          rs2.Open "SELECT SUM(CAST([mablaghgestall] AS money)) As rssum FROM aghsat WHERE (Groupme=" + Trim(Str(rs1.Fields!Groupme)) + ") AND (idbank=" + Trim(Str(rs1.Fields!idbank)) + ") AND (rad=" + Trim(Str(rs1.Fields!rad)) + ") And (Left(dateshamsi,4)='" + MaskEdBox7.Text + "')", db1
            If IsNull(rs2.Fields!rssum) = True Then
              Adodc5.Recordset.Fields!darvambahre = 0
            Else
              Adodc5.Recordset.Fields!darvambahre = rs2.Fields!rssum
            End If
          rs2.Close
          Adodc5.Recordset.Fields!darvamsum = Adodc5.Recordset.Fields!darvamasl - Adodc5.Recordset.Fields!darvambahre
          Adodc5.Recordset.Update
          q = q + 1
          rs1.MoveNext
        Loop Until rs1.EOF = True
      rs1.Close
    db1.Close
  End If
  
  Adodc5.ConnectionString = Form1.Text1.Text
  Adodc5.CommandType = adCmdUnknown
  Adodc5.RecordSource = "SELECT * FROM printvaz ORDER BY rad"
  Adodc5.Refresh
  
  Adodc5.Refresh
  DataGrid5.Refresh

  Adodc5.Refresh
  DataGrid5.Refresh
End If
End Sub

Private Sub Command11_Click()
If Adodc5.Recordset.RecordCount > 0 Then Form19.Show
End Sub

Private Sub Command2_Click()
Dim a1 As String, a2 As String, a3 As String, a4 As String
Dim S As String, w As Integer, e As String, r As String, t As String

w = 0
For q = 1 To 31
  MSChart1.Column = 1
  MSChart1.Row = q
  MSChart1.Data = 0
  MSChart1.Column = 2
  MSChart1.Row = q
  MSChart1.Data = 0
  MSChart1.RowLabel = q
Next q


MSChart1.Column = 1
db1.Open Form1.Text1.Text
  rs1.Open "SELECT count(rad) As rs1number FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Combo2.List(Combo1.ListIndex) + ") and (Left(dateshamsi,7)='" + MaskEdBox3.Text + "')", db1
    rs1number = rs1.Fields!rs1number
    Label1(8).Caption = rs1.Fields!rs1number
  rs1.Close
  If rs1number > 0 Then
    rs1.Open "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Combo2.List(Combo1.ListIndex) + ") and (Left(dateshamsi,7)='" + MaskEdBox3.Text + "')", db1
      rs1.MoveFirst
      Do
        MSChart1.Row = Val(Right(rs1.Fields!dateshamsi, 2))
        MSChart1.Data = Val(MSChart1.Data) + Val(rs1.Fields!Money) + Val(rs1.Fields!money1)
        rs1.MoveNext
      Loop Until rs1.EOF = True
    rs1.Close
  End If
db1.Close


MSChart1.Column = 2
db1.Open Form1.Text1.Text
  rs1.Open "SELECT count(rad) As rs1number FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Combo2.List(Combo1.ListIndex) + ") and (Left(datepardakhtall,7)='" + MaskEdBox3.Text + "') and (yesno='1')", db1
    rs1number = rs1.Fields!rs1number
    Label1(8).Caption = rs1.Fields!rs1number
  rs1.Close
  If rs1number > 0 Then
    rs1.Open "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Combo2.List(Combo1.ListIndex) + ") and (Left(datepardakhtall,7)='" + MaskEdBox3.Text + "') and (yesno='1')", db1
      rs1.MoveFirst
      Do
        MSChart1.Row = Val(Right(rs1.Fields!dateshamsi, 2))
        MSChart1.Data = Val(MSChart1.Data) + Val(rs1.Fields!Money) + Val(rs1.Fields!money1)
        rs1.MoveNext
      Loop Until rs1.EOF = True
    rs1.Close
  End If
db1.Close
End Sub

Private Sub Command3_Click()
Adodc1.Refresh
Adodc1.Refresh

Form15.Show
End Sub

Private Sub Command4_Click()
Dim moon(12) As String, moon1(12) As String
Dim a1 As String, a2 As String, a3 As String, a4 As String, a5 As String

db1.Open Form1.Text1.Text
  rs1.Open "DELETE FROM Printbanksmoon", db1
db1.Close

db1.Open Form1.Text1.Text
If Option1.Value = True Then
  rs1.Open "SELECT count(id) As rs1number FROM Bank", db1
    rs1number = rs1.Fields!rs1number
  rs1.Close
  If rs1number > 0 Then
    rs1.Open "SELECT * FROM Bank ORDER BY id ASC", db1
    Form5.Adodc1.Recordset.MoveFirst
      rs1.MoveFirst
      Do
        For q = 1 To 12
          moon1(q) = 0
        Next q
        For q = 2 To 13
          DataGrid3.Col = q
          DataGrid3.Columns.Item(q).Caption = Left(Amin.dateaminEzafeMoon(MaskEdBox4.Text + "/01", q - 2), 7)
          w = w + 1
          If w = 13 Then w = 1
          DoEvents
          
          rs2.Open "SELECT sum(CAST([money] AS money)) As rs1number,sum(CAST([money1] AS money)) As rs1number1 FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!ID)) + ") and (Left(dateshamsi,7)='" + DataGrid3.Columns.Item(q).Caption + "') and (yesno='0')", db1
            If IsNull(rs2.Fields!rs1number) = True Then
              moon1(q - 1) = 0
            Else
              moon1(q - 1) = Val(rs2.Fields!rs1number) + Val(rs2.Fields!rs1number1)
            End If
          rs2.Close
        Next q
        Adodc3.Recordset.AddNew
        Adodc3.Recordset.Fields!rad = rs1.Fields!ID
        Adodc3.Recordset.Fields!bankname = rs1.Fields!Name
        Adodc3.Recordset.Fields!m1 = Val(moon1(1))
        Adodc3.Recordset.Fields!m2 = Val(moon1(2))
        Adodc3.Recordset.Fields!m3 = Val(moon1(3))
        Adodc3.Recordset.Fields!m4 = Val(moon1(4))
        Adodc3.Recordset.Fields!m5 = Val(moon1(5))
        Adodc3.Recordset.Fields!m6 = Val(moon1(6))
        Adodc3.Recordset.Fields!m7 = Val(moon1(7))
        Adodc3.Recordset.Fields!m8 = Val(moon1(8))
        Adodc3.Recordset.Fields!m9 = Val(moon1(9))
        Adodc3.Recordset.Fields!m10 = Val(moon1(10))
        Adodc3.Recordset.Fields!m11 = Val(moon1(11))
        Adodc3.Recordset.Fields!m12 = Val(moon1(12))
        Adodc3.Recordset.Fields!a = "1"
        Adodc3.Recordset.Update
        rs1.MoveNext
      Loop Until rs1.EOF = True
    rs1.Close
  End If
Else
  a5 = Trim(Str(Val(Left(MaskEdBox4.Text, 4)) - 1)) + Right(MaskEdBox4.Text, 3)
  rs1.Open "SELECT count(id) As rs1number FROM Bank WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ")", db1
    rs1number = rs1.Fields!rs1number
  rs1.Close
  If rs1number > 0 Then
    rs1.Open "SELECT * FROM Bank Where (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") ORDER BY id ASC", db1
    Form5.Adodc1.Recordset.MoveFirst
      rs1.MoveFirst
      Do
        For q = 1 To 12
          moon1(q) = 0
        Next q
        For q = 2 To 13
          DataGrid3.Col = q
          DataGrid3.Columns.Item(q).Caption = Left(Amin.dateaminEzafeMoon(a5 + "/01", q - 2), 7)
          w = w + 1
          If w = 13 Then w = 1
          DoEvents
          
          rs2.Open "SELECT sum(CAST([money] AS money)) As rs1number,sum(CAST([money1] AS money)) As rs1number1 FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!ID)) + ") and (Left(dateshamsi,7)='" + DataGrid3.Columns.Item(q).Caption + "') and (yesno='1')", db1
            If IsNull(rs2.Fields!rs1number) = True Then
              moon1(q - 1) = 0
            Else
              moon1(q - 1) = Val(rs2.Fields!rs1number) + Val(rs2.Fields!rs1number1)
            End If
          rs2.Close
        Next q
        Adodc3.Recordset.AddNew
        Adodc3.Recordset.Fields!rad = rs1.Fields!ID
        Adodc3.Recordset.Fields!bankname = rs1.Fields!Name
        Adodc3.Recordset.Fields!m1 = Val(moon1(1))
        Adodc3.Recordset.Fields!m2 = Val(moon1(2))
        Adodc3.Recordset.Fields!m3 = Val(moon1(3))
        Adodc3.Recordset.Fields!m4 = Val(moon1(4))
        Adodc3.Recordset.Fields!m5 = Val(moon1(5))
        Adodc3.Recordset.Fields!m6 = Val(moon1(6))
        Adodc3.Recordset.Fields!m7 = Val(moon1(7))
        Adodc3.Recordset.Fields!m8 = Val(moon1(8))
        Adodc3.Recordset.Fields!m9 = Val(moon1(9))
        Adodc3.Recordset.Fields!m10 = Val(moon1(10))
        Adodc3.Recordset.Fields!m11 = Val(moon1(11))
        Adodc3.Recordset.Fields!m12 = Val(moon1(12))
        Adodc3.Recordset.Fields!a = "1"
        Adodc3.Recordset.Update
        rs1.MoveNext
      Loop Until rs1.EOF = True
    rs1.Close
  End If
End If
db1.Close
Adodc3.Refresh
Adodc3.Refresh

End Sub

Private Sub Command5_Click()
Form11.Show
End Sub

Private Sub Command6_Click()
Adodc2.Refresh
Adodc2.Refresh

Form13.Show
End Sub

Private Sub Command7_Click()
'  db1.Open Form1.Text1.Text
'    rs1.Open "DELETE FROM printvaz", db1
'  db1.Close
'
'  info(12) = 1
'  db1.Open Form1.Text1.Text
'    rs1.Open "SELECT * FROM Bank Where (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ")", db1
'      rs1.MoveFirst
'      Do
'        For q = 0 To 11
'          info(q) = 0
'        Next q
'
'        info(0) = rs1.Fields!ID
'        info(1) = rs1.Fields!Name
'        rs2.Open "SELECT count(rad) as rad1 FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!ID)) + ") and (tasvie='äÔÏå')", db1
'          numberrad = rs2.Fields!rad1
'        rs2.Close
'        If numberrad > 0 Then
'          rs2.Open "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs1.Fields!ID)) + ") and (tasvie='äÔÏå') ORDER BY rad ASC", db1
'            rs2.MoveFirst
'            Do
'              info(2) = Val(info(2)) + Val(rs2.Fields!moneyvamrial)
'              info(3) = Val(info(3)) + Val(rs2.Fields!allbahre)
'              info(4) = Val(info(4)) + Val(rs2.Fields!moneyvamrial) + Val(rs2.Fields!allbahre)
'
'              rs3.Open "SELECT count(rad) as rad1 FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs2.Fields!idbank)) + ") and (rad=" + Trim(Str(rs2.Fields!rad)) + ") and (yesno='1') ", db1
'                numberrad = rs3.Fields!rad1
'              rs3.Close
'              If numberrad > 0 Then
'                rs3.Open "SELECT * FROM aghsat WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(rs2.Fields!idbank)) + ") and (rad=" + Trim(Str(rs2.Fields!rad)) + ") and (yesno='1') ORDER BY rad2", db1
'                  rs3.Move numberrad - 1
'                  info(5) = Val(info(5)) + rs3.Fields!sumpardakhtasl
'                  info(6) = Val(info(6)) + rs3.Fields!sumpardakhtbahre
'                  info(7) = Val(info(7)) + rs3.Fields!sumpardakhtall
'                rs3.Close
'              End If
'              rs2.MoveNext
'            Loop Until rs2.EOF = True
'          rs2.Close
'        End If
'        info(8) = Val(info(2)) - Val(info(5))
'        info(9) = Val(info(3)) - Val(info(6))
'        info(10) = Val(info(4)) - Val(info(7))
'
'        info(11) = Val(rs1.Fields!saghf) - info(8)
'        Adodc1.Refresh
'        Adodc1.Recordset.AddNew
'        Adodc1.Recordset.Fields!rad = info(0)
'        Adodc1.Recordset.Fields!bankname = info(1)
'        Adodc1.Recordset.Fields!darvamasl = info(2)
'        Adodc1.Recordset.Fields!darvambahre = info(3)
'        Adodc1.Recordset.Fields!darvamsum = info(4)
'        Adodc1.Recordset.Fields!ghestvamasl = info(5)
'        Adodc1.Recordset.Fields!ghestvambahre = info(6)
'        Adodc1.Recordset.Fields!ghestvamsum = info(7)
'        Adodc1.Recordset.Fields!mondevamasl = info(8)
'        Adodc1.Recordset.Fields!mondevambahre = info(9)
'        Adodc1.Recordset.Fields!mondevamsum = info(10)
'        Adodc1.Recordset.Fields!saghf = info(11)
'        Adodc1.Recordset.Update
'        DataGrid1.Refresh
'        info(12) = Val(info(12)) + 1
'        rs1.MoveNext
'      Loop Until rs1.EOF = True
'    rs1.Close
'  db1.Close
'
'  Adodc1.CommandType = adCmdUnknown
'  Adodc1.RecordSource = "SELECT * FROM printvaz"
'  Adodc1.Refresh
'  DataGrid1.Refresh
'
'  Adodc1.Refresh
'  DataGrid1.Refresh
'
'  Adodc1.Recordset.Sort = "rad"

Adodc1.ConnectionString = Form1.Text1.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM View_vaz_banks"
Adodc1.Refresh
  
DataGrid1.Refresh
 
Adodc1.Recordset.Sort = "id"
End Sub

Private Sub Command8_Click()
If Adodc4.Recordset.RecordCount > 0 Then Form18.Show
End Sub

Private Sub Command9_Click()
If Combo6.ListIndex <> -1 Then
  Adodc4.ConnectionString = Form1.Text1.Text
  Adodc4.CommandType = adCmdUnknown
  Adodc4.RecordSource = "SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo5.List(Combo6.ListIndex))) + ") And (dateshamsi>='" + MaskEdBox5.Text + "') And (dateshamsi<='" + MaskEdBox6.Text + "')"
  Adodc4.Refresh
  db1.Open Form1.Text1.Text
    rs1.Open "SELECT Sum(moneyvamrial) As moneyvamrial12 FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo5.List(Combo6.ListIndex))) + ") And (dateshamsi>='" + MaskEdBox5.Text + "') And (dateshamsi<='" + MaskEdBox6.Text + "')", db1
      If IsNull(rs1.Fields!moneyvamrial12) = False Then
        Label3.Caption = Amin.moneyaminjoda(rs1.Fields!moneyvamrial12)
      Else
        Label3.Caption = 0
      End If
    rs1.Close
  db1.Close
End If
End Sub

Private Sub DataGrid1_DblClick()
DataGrid1.Col = 0
If Adodc1.Recordset.RecordCount > 0 Then
  If DataGrid1.Text <> "" Then
    db1.Open Form1.Text1.Text
      rs1.Open "DELETE FROM printvazriz", db1
    db1.Close
    q = DataGrid1.Text
    Form8.Show
    Form8.Text1.Text = q
    Me.Hide
  End If
End If
End Sub

Private Sub Form_Activate()
MaskEdBox1.Mask = "13##/##/##"
MaskEdBox1.PromptChar = "*"
MaskEdBox2.Mask = "13##/##/##"
MaskEdBox2.PromptChar = "*"
MaskEdBox3.Mask = "13##/##"
MaskEdBox3.PromptChar = "*"
MaskEdBox4.Mask = "13##/##"
MaskEdBox4.PromptChar = "*"
MaskEdBox5.Mask = "13##/##/##"
MaskEdBox5.PromptChar = "*"
MaskEdBox6.Mask = "13##/##/##"
MaskEdBox6.PromptChar = "*"
MaskEdBox7.Mask = "13##"
MaskEdBox7.PromptChar = "*"
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
Combo5.Clear
Combo6.Clear
Combo7.Clear
Combo8.Clear
Form5.Adodc1.ConnectionString = Form1.Text1.Text
Form5.Adodc1.CommandType = adCmdUnknown
Form5.Adodc1.RecordSource = "SELECT * FROM Bank Where (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") ORDER BY id"
Form5.Adodc1.Refresh
If Form5.Adodc1.Recordset.RecordCount > 0 Then
  Form5.Adodc1.Recordset.MoveFirst
  Do
    Combo2.AddItem Form5.Adodc1.Recordset.Fields!ID
    Combo1.AddItem Form5.Adodc1.Recordset.Fields!Name
    Combo4.AddItem Form5.Adodc1.Recordset.Fields!ID
    Combo3.AddItem Form5.Adodc1.Recordset.Fields!Name
    Combo5.AddItem Form5.Adodc1.Recordset.Fields!ID
    Combo6.AddItem Form5.Adodc1.Recordset.Fields!Name
    Combo7.AddItem Form5.Adodc1.Recordset.Fields!ID
    Combo8.AddItem Form5.Adodc1.Recordset.Fields!Name
    Form5.Adodc1.Recordset.MoveNext
  Loop Until Form5.Adodc1.Recordset.EOF = True
  Combo1.ListIndex = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub

Private Sub HScroll1_Change()
MSChart1.Left = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Call HScroll1_Change
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox2.SetFocus
End Sub

Private Sub MaskEdBox2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub MaskEdBox3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command2.SetFocus
End Sub

Private Sub MaskEdBox4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command4.SetFocus
End Sub

Private Sub MaskEdBox5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then MaskEdBox6.SetFocus
End Sub

Private Sub MaskEdBox6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command9.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
  Adodc1.ConnectionString = Form1.Text1.Text
  Adodc1.CommandType = adCmdUnknown
  Adodc1.RecordSource = "SELECT * FROM View_vaz_banks"
  Adodc1.Refresh
  
  DataGrid1.Refresh
  
  Adodc1.Recordset.Sort = "id"
End If
End Sub
