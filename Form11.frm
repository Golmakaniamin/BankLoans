VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form11 
   Caption         =   " ⁄Âœ«  »«‰ò Â« „«Â«‰Â"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim Report As New CrystalReport1
Report.Text4.SetText Form7.DataGrid3.Columns.Item(2).Caption
Report.Text5.SetText Form7.DataGrid3.Columns.Item(3).Caption
Report.Text6.SetText Form7.DataGrid3.Columns.Item(4).Caption
Report.Text7.SetText Form7.DataGrid3.Columns.Item(5).Caption
Report.Text8.SetText Form7.DataGrid3.Columns.Item(6).Caption
Report.Text9.SetText Form7.DataGrid3.Columns.Item(7).Caption
Report.Text10.SetText Form7.DataGrid3.Columns.Item(8).Caption
Report.Text11.SetText Form7.DataGrid3.Columns.Item(9).Caption
Report.Text12.SetText Form7.DataGrid3.Columns.Item(10).Caption
Report.Text13.SetText Form7.DataGrid3.Columns.Item(11).Caption
Report.Text14.SetText Form7.DataGrid3.Columns.Item(12).Caption
Report.Text15.SetText Form7.DataGrid3.Columns.Item(13).Caption
Report.Text22.SetText Amin.mil2shams(Date$)
Report.Text23.SetText Time$
Report.Text20.SetText Form1.Text2.Text
Screen.MousePointer = vbHourglass
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
