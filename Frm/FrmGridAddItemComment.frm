VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmGridAddItemComment 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ· „·«ÕŸ« "
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   7
      Top             =   3570
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   2415
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   990
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   8
      Top             =   3570
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   3480
      Y2              =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   660
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›:"
      Height          =   255
      Index           =   2
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   660
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›: "
      Height          =   255
      Index           =   1
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”ÿ—: "
      Height          =   255
      Index           =   0
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "FrmGridAddItemComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FG As VSFlex8UCtl.VSFlexGrid

Public LngRow As Long

Public LngCol As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String

    If Not FG Is Nothing Then
        FG.TextMatrix(LngRow, LngCol) = Trim$(Me.TxtComment.Text)
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub
Private Sub ChangeLang()

    Me.Caption = "Remarks"
lbl(0).Caption = "Line"
    lbl(1).Caption = "Code"
    Me.lbl(2).Caption = "Name"
    
    CmdOk.Caption = "Save"
    CmdCancel.Caption = "Close"
 
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

