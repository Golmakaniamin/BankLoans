VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form18 
   Caption         =   "ãÌãæÚ æÇã åÇ ÑíÇáí"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   LinkTopic       =   "Form18"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      lastProp        =   500
      _cx             =   10231
      _cy             =   12347
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim Report As New CrystalReport6

Screen.MousePointer = vbHourglass

'"SELECT * FROM vamtable WHERE (Groupme=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And (idbank=" + Trim(Str(Combo5.List(Combo6.ListIndex))) + ") And (dateshamsi>='" + MaskEdBox5.Text + "') And (dateshamsi<='" + MaskEdBox6.Text + "')"
'Report.RecordSelectionFormula = "({Ezam_p.date1} <> 'äÏÇÑÏ') AND ({Ezam_p.dore3} > 0)"

Report.RecordSelectionFormula = "({vamtable.Groupme}=" + Trim(Str(Form1.Combo2.List(Form1.Combo1.ListIndex))) + ") And ({vamtable.idbank}=" + Trim(Str(Form7.Combo5.List(Form7.Combo6.ListIndex))) + ") And ({vamtable.dateshamsi}>='" + Form7.MaskEdBox5.Text + "') And ({vamtable.dateshamsi}<='" + Form7.MaskEdBox6.Text + "')"
Report.Text3.SetText Amin.mil2shams(Date$)
Report.Text23.SetText Time$
Report.Text10.SetText Form7.Combo6.Text

CRViewer91.ReportSource = Report
CRViewer91.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub
