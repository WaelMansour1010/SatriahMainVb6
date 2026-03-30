VERSION 5.00
Begin VB.Form FrmShowReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÿ—Ìﬁ… «·⁄—÷"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   Icon            =   "FrmShowReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   720
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdShow 
      Caption         =   "› Õ «·„” ‰œ ··„—«Ã⁄Â"
      Height          =   375
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "«·ÿ»«⁄Â ⁄·Ï «·‘«‘…"
      Height          =   375
      Index           =   0
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmShowReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public reportno As Integer
Public NoteSerial As String

Private Sub Command1_Click()

End Sub

Private Sub CmdShow_Click(Index As Integer)
Select Case Index

Case 0
                If reportno = 200 Then
                   ShowGL_cc NoteSerial, , 200
                End If


Case 1


If reportno = 200 Then
            Unload FrmAccEditJournal
                        FrmAccEditJournal.Retrive NoteSerial
                        FrmAccEditJournal.show
                        FrmAccEditJournal.StrOldTransID = NoteSerial
                   
End If

End Select
Unload Me
End Sub
Private Sub ChangeLang()
    Me.Caption = "View Type..."
    CmdShow(0).Caption = "print On Screen"
    CmdShow(1).Caption = "Open Doc"
     
 
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    If Me.MDIChild = True Then
        Resize_Form Me
    Else
        CenterForm Me
    End If


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    


End Sub
