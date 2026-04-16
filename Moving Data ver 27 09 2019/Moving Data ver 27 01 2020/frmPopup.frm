VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Timer Timer1 
      Left            =   1590
      Top             =   930
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2835
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4485
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private RemainingTicks As Long

Public Sub ShowMessage(ByVal txt As String, Optional ByVal Seconds As Integer = 1)
    lblMsg.Caption = txt
    ' ‰Õ”» «·„œ… »«·‹ 1000 „·Ì À«‰Ì…
    RemainingTicks = Seconds
    Timer1.Interval = 100   ' ‰Œ·Ì «· «Ì„— Ì⁄œ ﬂ· À«‰Ì…
    Timer1.Enabled = True
    Me.Show vbModeless
End Sub

Private Sub Timer1_Timer()
    RemainingTicks = RemainingTicks - 1
    If RemainingTicks <= 0 Then
        Timer1.Enabled = False
        Unload Me
    End If
End Sub

