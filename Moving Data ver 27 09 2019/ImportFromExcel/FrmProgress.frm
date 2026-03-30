VERSION 5.00
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Begin VB.Form FrmProgress 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5820
      Top             =   570
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   1350
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin ImpulseAniLabel.ISAniLabel ISAniLabel19 
      Height          =   330
      Left            =   5910
      TabIndex        =   2
      Top             =   1740
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      AutoSize        =   -1  'True
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   14.25
      ForeColor       =   12582912
      BackColor       =   8421504
      Alignment       =   1
      Caption         =   "000"
      ForeFlashing    =   -1  'True
      ColorHover      =   12582912
      ColorAlternate  =   192
      UseColorAlternate=   -1  'True
      ImageCount      =   0
   End
   Begin AniGIFCtrl.AniGIF AniGIF2 
      Height          =   1605
      Left            =   0
      TabIndex        =   3
      Top             =   -30
      Width           =   1635
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "FrmProgress.frx":0000
      ExtendWidth     =   2884
      ExtendHeight    =   2831
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1740
      Width           =   2445
   End
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture _
                Lib "user32" () As Long
Private Sub AniGIF2_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    MoveForm
End Sub

Private Sub Form_Load()
    'CenterFor Me
    Me.ISAniLabel19.Font.Charset = 178
    ISAniLabel19.RightToLeft = True
    ISAniLabel19.Alignment = vbRightJustify
    ISAniLabel19.BackStyle = impTransparent
    ISAniLabel19.AutoSize = True

    
    Me.ISAniLabel19.Caption = "ÌÇÑì ÊÍãíá ÇáÈíÇäÇÊ íÑÌì ÇáÇäÊÙÇÑ  "
   

    ISAniLabel19.RedrawFull
    ISAniLabel19.Left = Me.PrgBar.Left + (Me.PrgBar.Width - Me.ISAniLabel19.Width)
    Me.lbl1.Left = 0
    'Me.lbl1.Width = Me.ISAniLabel19.left
    lbl1.BackStyle = vbTransparent
    Me.BackColor = vbWhite
End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    MoveForm
End Sub

Private Sub ISAniLabel19_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    MoveForm
End Sub

Private Sub lbl1_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    MoveForm
End Sub

Private Sub PrgBar_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    MoveForm
End Sub

Private Sub Timer1_Timer()
    Me.PrgBar.Value = FrmProgress.PrgBar.Value + 1
    Me.lbl1.Caption = IIf(Len(Me.lbl1.Caption) >= 10, String$(1, "-"), String(Len(Me.lbl1.Caption) + 1, "-"))

    If Me.PrgBar.Value = FrmProgress.PrgBar.Max Then
        Me.PrgBar.Value = 0
    End If

End Sub

Private Sub MoveForm()
    ReleaseCapture
    'SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
