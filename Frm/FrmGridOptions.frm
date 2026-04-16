VERSION 5.00
Begin VB.Form FrmGridOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ŒÌ«—«  ÃœÊ· ⁄—÷ «·√’‰«›"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13800
   Icon            =   "FrmGridOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   13800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   1095
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2520
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Õœœ «·‰Ê⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   8895
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕÊÌ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ”⁄—"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   0
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "›« Ê—…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   1
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "≈÷«›… «·”ÿ— «·ÃœÌœ  ·ﬁ«∆Ì«"
      Height          =   285
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CheckBox ChkAutoSize 
      Alignment       =   1  'Right Justify
      Caption         =   " ›⁄Ì· «· ÕÃÌ„ «· ·ﬁ«∆Ï"
      Height          =   285
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»«—ﬂÊœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   1
      Left            =   -480
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   4905
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   0
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   4905
   End
End
Attribute VB_Name = "FrmGridOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
