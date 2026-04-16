VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Splish 
   BorderStyle     =   0  'None
   Caption         =   "ALSATTARYAH GROUP"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   Icon            =   "Splish.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
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
      Interval        =   60000
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved Copyright © 2015 "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   240
      Picture         =   "Splish.frx":6852
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Byte Dynamic Integrated Software"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALSATTARYAH GROUP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   33
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   5520
      Left            =   0
      Picture         =   "Splish.frx":C9EA
      Stretch         =   -1  'True
      Top             =   -1560
      Width           =   8520
   End
End
Attribute VB_Name = "Splish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

Timer1.Enabled = True
Timer1.interval = 100
ProgressBar1.value = 100
Timer1_Timer
 End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label22_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
ProgressParChang
End Sub
Private Sub ProgressParChang()
If ProgressBar1.value = 100 Then
ProgressBar1.value = 0
Else
ProgressBar1.value = val(ProgressBar1.value) + val(1)
End If
Label1.Caption = ProgressBar1.value
End Sub
Private Sub Label1_Change()
ProgressBar1.Visible = True
If Label1.Caption = 100 Then
Timer1.interval = 0
Timer1.Enabled = False
ProgressBar1.Visible = False
Me.Visible = False
End If
End Sub
