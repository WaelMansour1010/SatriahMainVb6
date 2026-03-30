VERSION 5.00
Begin VB.Form FrmRegInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «· ”ÃÌ·"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "FrmRegInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "Œ—ÊÃ"
      Height          =   495
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3090
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5160
      X2              =   0
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   2625
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   2625
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   2625
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   2625
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   2625
   End
   Begin VB.Label LblCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   2625
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·:"
      Height          =   315
      Index           =   5
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·:"
      Height          =   315
      Index           =   4
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·:"
      Height          =   315
      Index           =   3
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·:"
      Height          =   315
      Index           =   2
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·:"
      Height          =   315
      Index           =   1
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·:"
      Height          =   315
      Index           =   0
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmRegInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub
