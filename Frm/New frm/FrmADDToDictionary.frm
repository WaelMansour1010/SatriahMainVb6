VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmADDToDictionary 
   Caption         =   "«÷«›… ··ﬁ«„Ê”"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3705
   Icon            =   "FrmADDToDictionary.frx":0000
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   3705
   Begin VB.TextBox Txename 
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Txtaname 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin ImpulseButton.ISButton XPButton301 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ButtonImage     =   "FrmADDToDictionary.frx":000C
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnsave 
      Height          =   345
      Left            =   870
      TabIndex        =   5
      Top             =   1080
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«÷«›…"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ButtonImage     =   "FrmADDToDictionary.frx":03A6
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label Label2 
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmADDToDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion
    clear_all Me
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnsave_Click()
Cn.Execute "insert into   edictionary (aname,ename) values ('" & Txtaname.Text & "','" & Txename.Text & "') "
MsgBox " „  «·«÷«›…"
End Sub

Private Sub XPButton301_Click()
Unload Me
End Sub
