VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form crLG_frm 
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
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
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "crLG_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CRViewer1.ReportSource = "select * from dblog where no_lg ='" & CetakLg_frm.Label1(8).Caption & "'"
End Sub
