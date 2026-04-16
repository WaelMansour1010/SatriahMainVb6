VERSION 5.00
Begin VB.Form INTRO 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "intro.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   15270
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
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9600
      Top             =   7560
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   3000
      Top             =   8400
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   12240
      TabIndex        =   4
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   13080
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      BorderWidth     =   5
      Height          =   735
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÔÑßÉ ÇáÓÊÇÑíÉ áÇäÙãÉ ÇáãÚáæãÇÊ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9600
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãÏÑÓÉ ÇáãäÇåÌ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   9600
      TabIndex        =   1
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ÏÎæá"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   0
      Top             =   9480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   6720
      Picture         =   "intro.frx":126E1
      Top             =   9000
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   6240
      Picture         =   "intro.frx":167A3
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   6
      Left            =   6240
      Picture         =   "intro.frx":19E6B
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   5
      Left            =   6360
      Picture         =   "intro.frx":1D533
      Top             =   6360
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   4
      Left            =   6360
      Picture         =   "intro.frx":20BFB
      Top             =   5520
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   3
      Left            =   6360
      Picture         =   "intro.frx":242C3
      Top             =   4680
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   2
      Left            =   6360
      Picture         =   "intro.frx":2798B
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   1
      Left            =   6240
      Picture         =   "intro.frx":2B053
      Top             =   3000
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   0
      Left            =   6360
      Picture         =   "intro.frx":2E71B
      Top             =   2160
      Width           =   2205
   End
   Begin VB.Image i1 
      Height          =   375
      Left            =   7080
      Picture         =   "intro.frx":31DE3
      Stretch         =   -1  'True
      Top             =   1785
      Width           =   1215
   End
End
Attribute VB_Name = "INTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numb As Integer

Private Sub Form_Load()
numb = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = True
End Sub

Private Sub Label1_Click()
login.Show
Timer2.Enabled = False
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = False
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label5_Click()
login.Show
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
If i1.Visible = True Then
i1.Visible = False

GoTo 10
End If
If numb > 7 And Image2.Visible = True Then
'Form2.Show

GoTo 10
End If
If numb > 7 Then
Image2.Visible = True
Label1.Visible = True
Timer2.Enabled = True
Timer1.Interval = 300
ElseIf numb <= 7 Then
If Image1(numb).Height > 15 Then
Image1(numb).Height = Image1(numb).Height - 120
Image1(numb).Top = Image1(numb).Top + 120
Else
Image1(numb).Visible = False
numb = numb + 1
End If
End If

10 End Sub

Private Sub Timer2_Timer()
If Label1.Visible = False Then Exit Sub
If Shape1.BorderColor = &H80000008 Then
Shape1.BorderColor = &H4000&
Else
Shape1.BorderColor = &H80000008
End If
End Sub
