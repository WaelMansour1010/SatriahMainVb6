VERSION 5.00
Begin VB.Form FrmItemTip 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4605
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3930
      Top             =   180
   End
End
Attribute VB_Name = "FrmItemTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    PutFormOnTop Me.hWnd, True
End Sub

Private Sub Timer1_Timer()
    Dim xP As POINTAPI
    Dim SngLeft As Single
    Dim SngTop As Single

    GetCursorPos xP
    SngLeft = xP.x * Screen.TwipsPerPixelX
    SngTop = xP.Y * Screen.TwipsPerPixelY

    Me.left = SngLeft + 100
    Me.top = SngTop + 100
End Sub
