VERSION 5.00
Begin VB.Form FrmNEWlOGIN 
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   4095
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "FrmNewLogin.frx":0000
      RightToLeft     =   -1  'True
      ScaleHeight     =   4035
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Command2 
         Caption         =   " ”ÃÌ·"
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   8040
         Top             =   2520
      End
      Begin VB.ComboBox CboInterface 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3240
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox dbname 
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox DbNamePath 
         Height          =   315
         Left            =   -960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " €ÌÌ— ﬁ«⁄œ… «·»Ì«‰« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õœœ ﬁ«⁄œ… «·»Ì«‰« "
         Height          =   255
         Left            =   6600
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Height          =   3975
         Left            =   0
         Top             =   0
         Width           =   8655
      End
   End
End
Attribute VB_Name = "FrmNEWlOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    save_login_info1 Trim(Me.DbNamePath.text), Trim(Me.dbname.text)
    AskForExit

End Sub

Public Function AskForExit() As Boolean
    End
End Function

Private Sub dbname_Click()
    DbNamePath.ListIndex = dbname.ListIndex

End Sub

Private Sub Form_Load()
    Open App.path & "\DB.txt" For Input As #1
    dbname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
                DbNamePath.AddItem (VarSet(0))
                dbname.AddItem (VarSet(1))
                            
            End If
        End If
    
    Loop

    Close #1

    Me.DbNamePath.text = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")
    Me.dbname.text = GetSetting("Byte_DBS", "Setting", "dbname", "«·ﬁ«⁄œ… «·«”«”Ì…")

End Sub

