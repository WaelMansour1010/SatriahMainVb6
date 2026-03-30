VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "REPORT~1.OCX"
Begin VB.Form FrmReportControl 
   Caption         =   "Form2"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9825
   Begin XtremeReportControl.ReportControl ReportControl1 
      Height          =   5985
      Left            =   660
      TabIndex        =   0
      Top             =   210
      Width           =   8355
      _Version        =   786432
      _ExtentX        =   14737
      _ExtentY        =   10557
      _StockProps     =   64
      ShowGroupBox    =   -1  'True
      ShowFooter      =   -1  'True
      RightToLeft     =   -1  'True
      RightToLeftReading=   -1  'True
      ShowFooterRows  =   -1  'True
   End
End
Attribute VB_Name = "FrmReportControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With Me.ReportControl1
        .Columns.Add 0, "ﬂÊœ «·’‰›", 100, True
        .Columns.Add 1, "«”„ «·’‰›", 100, True
        .Columns.Add 2, "ﬂÊœ «·„Ã„Ê⁄…", 100, True
        .Columns.Add 3, "«”„ «·„Ã„Ê⁄…", 100, True
        .Columns.Add 4, "ﬂ„Ì… «·’‰›", 100, True
    End With

End Sub

Private Sub Form_Resize()
    Dim x As ReportColumn
    On Error Resume Next
    Me.ReportControl1.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
    Set x = Me.ReportControl1.Columns.find(0)
    x.Caption = "»«Ì  ··»—„ÃÌ« "
End Sub
