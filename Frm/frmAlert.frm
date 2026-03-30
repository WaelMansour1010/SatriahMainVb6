VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2490
      Top             =   1740
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   720
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Label lblAlert 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Message"
         Height          =   825
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1020
         Width           =   2145
      End
   End
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2520
      Top             =   1230
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   11
      Left            =   4200
      Picture         =   "frmAlert.frx":0000
      Top             =   0
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   10
      Left            =   2010
      Picture         =   "frmAlert.frx":3D02
      Top             =   870
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   9
      Left            =   900
      Picture         =   "frmAlert.frx":7A06
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   8
      Left            =   900
      Picture         =   "frmAlert.frx":B70A
      Top             =   2010
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   7
      Left            =   3120
      Picture         =   "frmAlert.frx":F40E
      Top             =   870
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   6
      Left            =   4260
      Picture         =   "frmAlert.frx":13112
      Top             =   870
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   5
      Left            =   2040
      Picture         =   "frmAlert.frx":16E16
      Top             =   2010
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   4
      Left            =   3150
      Picture         =   "frmAlert.frx":1AB1A
      Top             =   2010
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   3
      Left            =   4260
      Picture         =   "frmAlert.frx":1E81E
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   2
      Left            =   4260
      Picture         =   "frmAlert.frx":22522
      Top             =   2010
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   1
      Left            =   3150
      Picture         =   "frmAlert.frx":26226
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Image Img 
      Height          =   1080
      Index           =   0
      Left            =   2040
      Picture         =   "frmAlert.frx":29F2A
      Top             =   3120
      Width           =   1080
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private AlertIndex As Long
Dim m_AlertIndex As Integer


Private Sub Form_Unload(Cancel As Integer)
'AlertCountFree = Me.AlertIndex
'm_IndexCollection.Item(AlertCountFree) = 0
'AlertCount = AlertCount - 1
End Sub

Private Sub lblAlert_Click()
    ' When user clicked the alertbox
    MsgBox "ÓæÝ íÊã ÚÑÖ ÊÞÑíÑ Çæ ÔÇÔÉ ÊÚÑÖ" & Chr(13) & _
    "ÍÇáÉ ÇáÓåã Çæ ÇáÚãáíÉ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub lblAlert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show as hyperlink
    If lblAlert.FontUnderline = False Then
        lblAlert.FontUnderline = True
        lblAlert.ForeColor = RGB(0, 0, 255)
    End If
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show text
    If lblAlert.FontUnderline = True Then
        lblAlert.FontUnderline = False
        lblAlert.ForeColor = &H0
    End If
End Sub

Private Sub picBackground_Resize()
'lblAlert.Move picBackground.ScaleLeft, picBackground.ScaleTop, _
    picBackground.ScaleWidth, picBackground.ScaleHeight
End Sub

Private Sub tmrAlert_Timer()
    ' Alert was displayed, now close it
'    tmrAlert.Enabled = False
'    tmrClose.Enabled = True
End Sub

Private Sub tmrClose_Timer()
    Dim curHeight As Long
    curHeight = Me.Height
    If curHeight > 120 Then
        Me.Height = curHeight - 30
        Me.top = Me.top + 30
    Else
        ' Close form
'        If AlertCount = Me.AlertIndex Then
'            AlertCount = 0
'        Else
'            AlertCount = AlertCount - 1
'        End If
        AlertCountFree = Me.AlertIndex
        Unload Me
    End If
End Sub

Private Sub tmrOpen_Timer()
Dim curHeight As Long
Dim newHeight As Long
curHeight = Me.Height
If curHeight < picBackground.Height + lngScaleY Then
    newHeight = curHeight + 30
    If newHeight > picBackground.Height + lngScaleY Then
        newHeight = picBackground.Height + lngScaleY
    End If
    Me.Height = Me.Height + (newHeight - curHeight)
    Me.top = Me.top - (newHeight - curHeight)
Else
    tmrOpen.Enabled = False
'    tmrAlert.Enabled = True
End If
End Sub



Public Property Get AlertIndex() As Integer
AlertIndex = m_AlertIndex
End Property

Public Property Let AlertIndex(ByVal vNewValue As Integer)
m_AlertIndex = vNewValue
End Property

