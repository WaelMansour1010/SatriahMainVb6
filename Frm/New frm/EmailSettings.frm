VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmailSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Settings"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "EmailSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "HTML Mail"
      Height          =   495
      Left            =   6675
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3240
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Body"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   75
      TabIndex        =   10
      Top             =   1350
      Width           =   7890
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   7320
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1125
         TabIndex        =   24
         Text            =   "a.s@sattaryah.com"
         Top             =   700
         Width           =   2715
      End
      Begin VB.TextBox txtMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "EmailSettings.frx":000C
         Top             =   1500
         Width           =   6615
      End
      Begin VB.TextBox txtAttach 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5025
         TabIndex        =   15
         Top             =   650
         Width           =   2115
      End
      Begin VB.TextBox txtFromEmail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5025
         TabIndex        =   14
         Text            =   "info@sattaryah.com"
         Top             =   225
         Width           =   2715
      End
      Begin VB.TextBox txtSubject 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         TabIndex        =   13
         Text            =   "ReMinder Test"
         Top             =   1075
         Width           =   6615
      End
      Begin VB.TextBox txtFromName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         TabIndex        =   12
         Text            =   "Dynamic ERP"
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Attachement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   4050
         TabIndex        =   22
         Top             =   675
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   21
         Top             =   705
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   20
         Top             =   1100
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   1500
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   17
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SMTP Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   7890
      Begin VB.CheckBox chkSSL 
         Alignment       =   1  'Right Justify
         Caption         =   "Req. SSL"
         Height          =   315
         Left            =   2475
         TabIndex        =   11
         Top             =   675
         Width           =   1065
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   687
         TabIndex        =   5
         Text            =   "mail.sattaryah.com"
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   687
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "25"
         Top             =   690
         Width           =   600
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3321
         TabIndex        =   3
         Text            =   "a.s@sattaryah.com"
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5925
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "spamkiller"
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5178
         TabIndex        =   9
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2544
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   7
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   6675
      TabIndex        =   0
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   75
      TabIndex        =   23
      Top             =   4500
      Width           =   6390
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "EmailSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : SendMail
' Author    : coolcurrent4u
' Date      : 4/19/2011
' Purpose   : sends email using the cdo namespace
' TODO      : check for attachment existence before passing it
'           : Pass only number to port textfield or else it will throw an error
' Questions : Please ask in vbforums.com
'---------------------------------------------------------------------------------------
'



Private Sub cmdSend_Click()
    
    Dim RetVal          As String
    Dim objControl      As Control
    'Validate first
    For Each objControl In Me.Controls
        If TypeOf objControl Is TextBox Then
            If Trim$(objControl.Text) = vbNullString And LCase$(objControl.Name) <> "txtattach" Then
                Label2.Caption = "Error: All fields are required!"
                Exit Sub
            End If
        End If
    Next
    
    'Send
    Frame1.Enabled = False
    Frame2.Enabled = False
    cmdSend.Enabled = False
    Label2.Caption = "Sending..."
    RetVal = SendMail(Trim$(txtTo.Text), _
        Trim$(txtSubject.Text), _
        Trim$(txtFromName.Text) & "<" & Trim$(txtFromEmail.Text) & ">", _
        Trim$(txtMsg.Text), _
        Trim$(txtServer.Text), _
        CInt(Trim$(txtPort.Text)), _
        Trim$(txtUsername.Text), _
        Trim$(txtPassword.Text), _
        Trim$(txtAttach.Text), _
        CBool(chkSSL.value))
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdSend.Enabled = True
    Label2.Caption = IIf(RetVal = "ok", "Message sent!", RetVal)
    
End Sub



Private Sub Command1_Click()
CD1.ShowOpen
txtAttach.Text = CD1.filename

End Sub

Private Sub Command2_Click()
strBody = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf
strBody = strBody & "<html>" & vbCrLf
strBody = strBody & "<head>" & vbCrLf
 strBody = strBody & " <title>Internal mailer example</title>" & vbCrLf
strBody = strBody & " <meta http-equiv=Content-Type content=""text/html; charset=iso-8859-1"">" & vbCrLf
strBody = strBody & "</head>" & vbCrLf
strBody = strBody & "<body bgcolor=""#FFFFCC"">" & vbCrLf
strBody = strBody & " <h2> This is what your internal mail will look like</h2>" & vbCrLf
strBody = strBody & " <p>" & vbCrLf
strBody = strBody & " This message was sent from a sample at" & vbCrLf
strBody = strBody & " <a href=""http://www.hello-world.com"">hello world</a>." & vbCrLf
strBody = strBody & " It is used to show people how to send HTML" & vbCrLf
strBody = strBody & " formatted email if you want this kind of format rther than plain text" & vbCrLf
strBody = strBody & " This is what your internal to external mail ," & vbCrLf
strBody = strBody & "will look like, if you have any questions" & vbCrLf
strBody = strBody & " get in touch with us at blah balh balh." & vbCrLf
strBody = strBody & " <strong>" & vbCrLf
strBody = strBody & "mesage generated by the MSMQ and our business recoreds." & vbCrLf
strBody = strBody & " </strong>" & vbCrLf
strBody = strBody & " </p>" & vbCrLf
strBody = strBody & " <font size=""-1"">" & vbCrLf
strBody = strBody & " <p>Please address all concerns to hello wrold@everywhere.com.</p>" & vbCrLf
strBody = strBody & " <p>This message was sent to: " & strMailer & "</p>" & vbCrLf
strBody = strBody & " </font>" & vbCrLf
strBody = strBody & "</body>" & vbCrLf
strBody = strBody & "</html>" & vbCrLf
txtMsg.Text = strBody
cmdSend_Click
txtMsg.Text = ""
End Sub

