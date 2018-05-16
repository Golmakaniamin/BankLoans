VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form21 
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form21"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   10755
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   5085
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9285
      lastProp        =   500
      _cx             =   16378
      _cy             =   8969
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   495
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Screen.MousePointer = vbHourglass

Select Case Label1.Caption

  Case 8
    Dim Report8 As New CrystalReport8
    
    Me.Caption = "—Ì“ »œÂÌ"
    Report8.Text21.SetText "—Ì“ »œÂÌ"
    
    Report8.Text3.SetText Amin.mil2shams(Date$)
    Report8.Text1.SetText Time$
    
    Report8.ParameterFields(1).AddCurrentValue Val(Form1.Combo2.List(Form1.Combo1.ListIndex))
    Report8.ParameterFields(2).AddCurrentValue Form20.MaskEdBox1.Text
    Report8.ParameterFields(3).AddCurrentValue Form20.MaskEdBox2.Text
    Report8.ParameterFields(4).AddCurrentValue Form20.Combo2.List(Form20.Combo1.ListIndex)
    
    Report8.PrinterSetup (Me.hwnd)
    CRViewer91.ReportSource = Report8
    
  Case 9
    Dim Report9 As New CrystalReport9
    
    Me.Caption = "—Ì“ »œÂÌ  «—ÌŒ"
    Report9.Text21.SetText "—Ì“ »œÂÌ  «—ÌŒ"
    
    Report9.Text3.SetText Amin.mil2shams(Date$)
    Report9.Text10.SetText Time$
    Report9.ParameterFields(1).AddCurrentValue Val(Form1.Combo2.List(Form1.Combo1.ListIndex))
    Report9.ParameterFields(2).AddCurrentValue Form20.MaskEdBox3.Text
    Report9.ParameterFields(3).AddCurrentValue Form20.MaskEdBox4.Text
    Report9.ParameterFields(4).AddCurrentValue Form20.Combo5.List(Form20.Combo3.ListIndex)
    Report9.PrinterSetup (Me.hwnd)
    
    CRViewer91.ReportSource = Report9
    
End Select
CRViewer91.ViewReport
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub

