VERSION 5.00
Begin VB.Form InfoBar 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Interval        =   5000
      Left            =   600
      Top             =   240
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   0
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·„ÊŸ›"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„ÊŸ›"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LblOprname 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label LblTermName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblprojectName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„·Ì…"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·»‰œ"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„‘—Ê⁄"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label LblEmpName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label LblEmpCode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label LblEmpID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„ÊŸ›"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "InfoBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
