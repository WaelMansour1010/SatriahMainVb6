VERSION 5.00
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmExpire 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈‰ Â«¡ «·’·«ÕÌ… (‰”Œ…  Ã—Ì»Ì…)"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   Icon            =   "FrmExpire.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PctExpire 
      BorderStyle     =   0  'None
      Height          =   5805
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   5805
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      Begin ImpulseButton.ISButton cmdok 
         Height          =   315
         Left            =   1755
         TabIndex        =   1
         Top             =   4920
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "„Ê«›ﬁ"
         BackColor       =   13276201
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpire.frx":038A
         ButtonImageDisabled=   "FrmExpire.frx":0724
         ColorButton     =   13276201
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledText=   16777215
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseAniLabel.ISAniLabel LblLink 
         Height          =   285
         Left            =   1530
         TabIndex        =   2
         Top             =   4410
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         ActiveUnderline =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   16777215
         MousePointer    =   99
         MouseIcon       =   "FrmExpire.frx":0ABE
         BackColor       =   16777215
         Caption         =   "WWW.bisegypt.com"
         ColorHover      =   16711680
         ImageCount      =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1230
         Picture         =   "FrmExpire.frx":0C20
         Top             =   4410
         Width           =   240
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "bisegypt@yahoo.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   930
         TabIndex        =   7
         Top             =   4005
         Width           =   2925
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‘—ﬂÏ »«Ì  ··»—„ÃÌ«  «·„Õ«”»Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   750
         TabIndex        =   6
         Top             =   3120
         Width           =   3195
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê“⁄ «·„⁄ „œ ›Ì Ã„ÂÊ—Ì… „’— «·⁄—»Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   1050
         TabIndex        =   5
         Top             =   2820
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " ·Ì›Ê‰ : 0226210707 - 0226210706"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   1140
         TabIndex        =   4
         Top             =   3420
         Width           =   2625
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " : «·»—Ìœ «·≈·ﬂ —Ê‰Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   2340
         TabIndex        =   3
         Top             =   3720
         Width           =   1425
      End
   End
End
Attribute VB_Name = "FrmExpire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me

    If Dir(App.path & "\Garphics\Expire.bmp") <> "" Then
        PctExpire.Picture = LoadPicture(App.path & "\Garphics\Expire.bmp")
    End If

End Sub

Private Sub LblLink_Click()
    OpenWebSite
End Sub

