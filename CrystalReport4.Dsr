VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrystalReport4 
   ClientHeight    =   14565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23865
   OleObjectBlob   =   "CrystalReport4.dsx":0000
End
Attribute VB_Name = "CrystalReport4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Initialize()
For q = 1 To Database.Tables.Count
  Database.Tables(q).ConnectBufferString = Form1.Text1.Text
Next q
End Sub

