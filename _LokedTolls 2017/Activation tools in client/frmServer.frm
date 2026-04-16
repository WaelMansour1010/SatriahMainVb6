VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmActivation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " SQL  License Activaton"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   11040
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
   Begin VB.TextBox TxtNoOFUsers 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   960
      Width           =   375
   End
   Begin VB.CheckBox VbEcnomy 
      Alignment       =   1  'Right Justify
      Caption         =   "Lite"
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox PaysecondIns 
      Alignment       =   1  'Right Justify
      Caption         =   "INS"
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   71
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "$"
      TabIndex        =   68
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox TXTTechnicalId 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "$"
      TabIndex        =   66
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Advanced"
      Height          =   375
      Left            =   5280
      TabIndex        =   65
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   10200
      TabIndex        =   61
      Top             =   3480
      Width           =   9615
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«Œ—  «—ÌŒ ’Ì«‰…"
         Height          =   255
         Left            =   5160
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "  «—ÌŒ «” Õﬁ«ﬁ «·ﬁ”ÿ «· «·Ì"
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„” Œœ„Ì‰"
         Height          =   255
         Left            =   4560
         TabIndex        =   62
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«·„ÊœÌÊ·« "
      Enabled         =   0   'False
      Height          =   3735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3360
      Width           =   9615
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   50
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   49
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   48
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   47
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   46
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   45
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   44
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   43
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   42
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   41
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   40
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Õ”«»« "
         Height          =   255
         Index           =   0
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„⁄«„·«  «·„«·Ì…"
         Height          =   255
         Index           =   1
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«· Õ·Ì· «·„«·Ì"
         Height          =   255
         Index           =   2
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«’Ê· «·À«Ì …"
         Height          =   255
         Index           =   3
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "‘ „"
         Height          =   255
         Index           =   4
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„Œ“Ê‰"
         Height          =   255
         Index           =   5
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„‘ —Ì« "
         Height          =   255
         Index           =   6
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«· ”ÊÌﬁ"
         Height          =   255
         Index           =   7
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„»Ì⁄« "
         Height          =   255
         Index           =   8
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‘Õ‰"
         Height          =   255
         Index           =   9
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "‰ﬁ«ÿ «·»Ì⁄"
         Height          =   255
         Index           =   10
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«‰ «Ã"
         Height          =   255
         Index           =   11
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„—«ﬁ»… «·ÃÊœ…"
         Height          =   255
         Index           =   12
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„‘«—Ì⁄"
         Height          =   255
         Index           =   13
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‰ﬁ·Ì« "
         Height          =   255
         Index           =   14
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "’Ì«‰… «·„⁄œ« /«·”Ì«—« "
         Height          =   255
         Index           =   15
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·’Ì«‰… «·⁄«„…"
         Height          =   255
         Index           =   16
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‰ﬁ· «·„œ—”Ì"
         Height          =   255
         Index           =   17
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«—‘Ì›"
         Height          =   255
         Index           =   18
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ«—” Ê«·„⁄«Âœ «· ⁄·Ì„Ì…"
         Height          =   255
         Index           =   19
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«”Â„"
         Height          =   255
         Index           =   20
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«œ«—… «·«„·«ﬂ"
         Height          =   255
         Index           =   21
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«· ﬁ«—Ì—"
         Height          =   255
         Index           =   22
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„œÌ— «·‰Ÿ«„"
         Height          =   255
         Index           =   23
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
         Height          =   255
         Index           =   24
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«œÊ«  «·›"
         Height          =   255
         Index           =   25
         Left            =   -360
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«· ÿÊÌ—"
         Height          =   255
         Index           =   26
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«· ŒÿÌÿ"
         Height          =   255
         Index           =   27
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkselectall 
         Alignment       =   1  'Right Justify
         Caption         =   "«Œ Ì«— «·ﬂ·"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   28
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkmai 
         Alignment       =   1  'Right Justify
         Caption         =   "«·»—‰«„Ã «·⁄«„"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   29
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‘ﬁﬁ «·›‰œﬁÌ…"
         Height          =   255
         Index           =   28
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„“«—⁄ «·œÊ«Ã‰"
         Height          =   255
         Index           =   29
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "Ê—‘ «·œÂ» Ê«·«·„«”"
         Height          =   255
         Index           =   30
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„⁄«„·«  «·»‰ﬂÌ…"
         Height          =   255
         Index           =   31
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‘∆Ê‰ «·«œ«—Ì…"
         Height          =   255
         Index           =   32
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·”„«Â„«  «·⁄ﬁ«—Ì…"
         Height          =   255
         Index           =   33
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„»Ì⁄«  «· ﬁ”Ìÿ"
         Height          =   255
         Index           =   34
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«œ«—… «·„’«⁄œ"
         Height          =   255
         Index           =   35
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ÕÃ Ê «·⁄„—…"
         Height          =   255
         Index           =   36
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„ƒ‘—«  «·ÕÌ…"
         Height          =   255
         Index           =   37
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   38
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕÃÊ“"
         Height          =   255
         Index           =   39
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Activate"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox SQlTxt 
      Height          =   2175
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtDexrypted 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   9360
      Width           =   6975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TxtLicense 
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1680
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ›⁄Ì·"
      Height          =   495
      Left            =   18120
      TabIndex        =   8
      Top             =   4800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "GetCode By"
      Height          =   1695
      Left            =   14160
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton OptActtype 
         Caption         =   "Direct"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton OptActtype 
         Caption         =   "Email"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton OptActtype 
         Caption         =   "Sms"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtCode 
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ›⁄Ì·"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   8520
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker LockedDate 
      Height          =   345
      Left            =   5640
      TabIndex        =   70
      Top             =   840
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   76611585
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker Alarm_start 
      Height          =   345
      Left            =   3720
      TabIndex        =   72
      Top             =   840
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   76611585
      CurrentDate     =   38784
   End
   Begin VB.Image Image1 
      Height          =   3075
      Left            =   8640
      Picture         =   "frmServer.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2490
   End
   Begin VB.Label Label8 
      Caption         =   "Advance Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   69
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Activate Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   67
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   3240
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Activation Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label License 
      Caption         =   "License"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   18120
      TabIndex        =   7
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label lbl 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim publickey1 As String
Private Declare Function SendMessageAsLong Lib "user32" _
     Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long
Private Type tGUID
   l1 As Long
   l2 As Long
   l3 As Long
   l4 As Long
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" ( _
      lpGuid As tGUID _
   ) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" ( _
      lpGuid As tGUID, _
      ByVal lpString As String, _
      ByVal cbBytes As Integer _
   ) As Integer
Public Function GetNetworkConnectionMACAddress() As String

' Return the currently used network adapter's MAC address

' Syntax
'
' GetNetworkConnectionMACAddress()

    Dim oWMIService As Object
    Dim vAdapters As Variant
    Dim oAdapter As Object
    Dim lIndex As Long
    Dim lMatchIndex As Long
    Dim vResult As Variant
    
    ' Adapters are pulled from the Windows Management Instrumentation database
    ' The currently used adapter has a MAC address and an IP address that is not 0.0.0.0
    Set oWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    Set vAdapters = oWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each oAdapter In vAdapters
        If Not IsNull(oAdapter.MACAddress) And IsArray(oAdapter.IPAddress) Then
            lMatchIndex = -1
            For lIndex = 0 To UBound(oAdapter.IPAddress)
                If Not oAdapter.IPAddress(lIndex) = "0.0.0.0" Then
                    lMatchIndex = lIndex
                    Exit For
                End If
            Next lIndex
            If Not lMatchIndex < 0 Then
                GetNetworkConnectionMACAddress = oAdapter.MACAddress
            End If
        End If
   Next

End Function

 


Public Function CreateGUID() As String

' Create and return a unique GUID string.

   Dim GUID As tGUID
   Dim Temp As String
   Dim Result As Long
   Dim Length As Long
   
   Result = CoCreateGuid(GUID)
   If (Result = 0) Then
      Temp = StrConv(String(38, Chr(0)), vbUnicode)
      Length = StringFromGUID2(GUID, Temp, Len(Temp))
      Temp = StrConv(Temp, vbFromUnicode)
      If (Length > 0) Then
         If (Left(Temp, 1) = "{") Then Temp = Right(Temp, Len(Temp) - 1)
         If (Right(Temp, 1) = "}") Then Temp = Left(Temp, Len(Temp) - 1)
         Length = InStr(Temp, "-")
         Do While (Length <> 0)
            Temp = Left(Temp, Length - 1) & Right(Temp, Len(Temp) - Length)
            Length = InStr(Temp, "-")
         Loop
      Else
         Temp = ""
      End If
   End If
   CreateGUID = Temp

End Function
Function URLEncode(ByVal str As String) As String
    Dim intLen As Integer
    Dim X As Integer
    Dim curChar As Long
    Dim newStr As String

    intLen = Len(str)
    newStr = ""

    For X = 1 To intLen
        curChar = Asc(Mid$(str, X, 1))
          
        If (curChar < 48 Or curChar > 57) And (curChar < 65 Or curChar > 90) And (curChar < 97 Or curChar > 122) Then
            newStr = newStr & "%" & Hex(curChar)
        Else
            newStr = newStr & Chr(curChar)
        End If

    Next X
              
    URLEncode = newStr
End Function


Public Sub SendMessage(Optional msgstr As String = "", _
                       Optional Numbers As String = "")
    Dim t As String

    If msgstr = "" Then
        msgstr = txtMessage.Text
    End If

    If Numbers = "" Then
        Numbers = txtNumbers.Text
    End If

    ''t = send(UserName, URLEncode(Password), ConvertToUnicode(ConvertString(txtMessage.Text)), txtSender.Text, txtNumbers.Text)
    't = Send("966550015230 ", URLEncode("aljazeera10"), ConvertToUnicode(msgstr), txtSender.Text, Numbers)
 
    If msgstr = "" Then
        ShowResult (t)
    Else
        ShowResult t, 1
    End If

End Sub
Private Sub ShowResult(val As String, _
                       Optional outme As Integer = 0)

    If outme <> 0 Then Exit Sub

    Select Case val

        Case "1": MsgBox ("·ﬁœ  „   ⁄„·Ì… «—”«· «·—”«·…  »‰Ã«Õ") 'sent

        Case "2": MsgBox ("≈‰ —’Ìœﬂ ·œÏ „Ê»«Ì·Ì ﬁœ ≈‰ ÂÏ Ê·„ Ì⁄œ »Â √Ì —”«∆·. (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance = 0

        Case "3": MsgBox ("≈‰ —’Ìœﬂ «·Õ«·Ì ·« Ìﬂ›Ì ·≈ „«„ ⁄„·Ì… «·≈—”«·. (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance  not  enough"

        Case "4": MsgBox ("≈‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ··œŒÊ· ≈·Ï Õ”«» «·—”«∆· €Ì— ’ÕÌÕ ( √ﬂœ „‰ √‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ÂÊ ‰›”Â «·–Ì  ” Œœ„Â ⁄‰œ œŒÊ·ﬂ ≈·Ï „Êﬁ⁄ „Ê»«Ì·Ì)") 'mobile not found

        Case "5": MsgBox ("Â‰«ﬂ Œÿ√ ›Ì ﬂ·„… «·„—Ê— ( √ﬂœ „‰ √‰ ﬂ·„… «·„—Ê— «· Ì  „ ≈” Œœ«„Â« ÂÌ ‰›”Â« «· Ì  ” Œœ„Â« ⁄‰œ œŒÊ·ﬂ „Êﬁ⁄ „Ê»«Ì·Ì,≈–« ‰”Ì  ﬂ·„… «·„—Ê— ≈÷€ÿ ⁄·Ï —«»ÿ ‰”Ì  ﬂ·„… «·„—Ê— · ’·ﬂ —”«·… ⁄·Ï ÃÊ«·ﬂ »—ﬁ„ «·„—Ê— «·Œ«’ »ﬂ)") 'password error

        Case "6": MsgBox ("≈‰ ’›Õ… «·≈—”«· ·« ÃÌ» ›Ì «·Êﬁ  «·Õ«·Ì (ﬁœ ÌﬂÊ‰ Â‰«ﬂ ÿ·» ﬂ»Ì— ⁄·Ï «·’›Õ… √Ê  Êﬁ› „ƒﬁ  ··’›Õ… ›ﬁÿ Õ«Ê· „—… √Œ—Ï √Ê  Ê«’· „⁄ «·œ⁄„ «·›‰Ì ≈–« ≈” „— «·Œÿ√)") 'page not response try send again

        Case "12": MsgBox ("≈‰ Õ”«»ﬂ »Õ«Ã… ≈·Ï  ÕœÌÀ Ì—ÃÏ „—«Ã⁄… «·œ⁄„ «·›‰Ì")

        Case "13": MsgBox ("≈‰ ≈”„ «·„—”· «·–Ì ≈” Œœ„ Â ›Ì Â–Â «·—”«·… ·„ Ì „ ﬁ»Ê·Â. (Ì—ÃÏ ≈—”«· «·—”«·… »≈”„ „—”· ¬Œ— √Ê  ⁄—Ì› ≈”„ «·„—”· ·œÏ „Ê»«Ì·Ì)") 'sender not accept

        Case "14": MsgBox "≈‰ ≈”„ «·„—”· «·–Ì ≈” Œœ„ Â €Ì— „⁄—› ·œÏ „Ê»«Ì·Ì. (Ì„ﬂ‰ﬂ  ⁄—Ì› ≈”„ «·„—”· „‰ Œ·«· ’›Õ… ≈÷«›… ≈”„ „—”·)" 'sender name not activated

        Case "15": MsgBox "ÌÊÃœ —ﬁ„ ÃÊ«· Œ«ÿ∆ ›Ì «·√—ﬁ«„ «· Ì ﬁ„  »«·≈—”«· ·Â«. ( √ﬂœ „‰ ’Õ… «·√—ﬁ«„ «· Ì  —Ìœ «·≈—”«· ·Â« Ê√‰Â« »«·’Ì€… «·œÊ·Ì…)"

        Case "16": MsgBox "«·—”«·… «· Ì ﬁ„  »≈—”«·Â« ·«  Õ ÊÌ ⁄·Ï ≈”„ „—”·. (√œŒ· ≈”„ „—”· ⁄‰œ ≈—”«·ﬂ «·—”«·…)"

        Case "17": MsgBox "·„ Ì „ «—”«· ‰’ «·—”«·…. «·—Ã«¡ «· √ﬂœ „‰ «—”«· ‰’ «·—”«·… Ê«· √ﬂœ „‰  ÕÊÌ· «·—”«·… «·Ï ÌÊ‰Ì ﬂÊœ («·—Ã«¡ «· √ﬂœ „‰ «” Œœ«„ «·œ«·… ConvertToUnicode)"

        Case "-1": MsgBox "·„ Ì „ «· Ê«’· „⁄ Œ«œ„ (Server) «·≈—”«· „Ê»«Ì·Ì »‰Ã«Õ. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"

        Case "-2": MsgBox "·„ Ì „ «·—»ÿ „⁄ ﬁ«⁄œ… «·»Ì«‰«  (Database) «· Ì  Õ ÊÌ ⁄·Ï Õ”«»ﬂ Ê»Ì«‰« ﬂ ·œÏ „Ê»«Ì·Ì. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
    
        Case Else: MsgBox (val)
    End Select

End Sub

Private Sub Command1_Click()
TxtCode = CreateGUID
'SendMessage TxtCode, "966541793243"


End Sub
Public Function CryptRC4(sText As String, sKey As String) As String
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap     As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lIdx        As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    lI = 0
    lJ = 0
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
    Next
End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
End Function

Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long

    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
    Next
End Function

Public Function FromHexDump(sText As String) As String
    Dim lIdx            As Long

    For lIdx = 1 To Len(sText) Step 2
        FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid(sText, lIdx, 2)))
    Next
End Function
Private Sub Command2_Click()
    
 
Dim myWMI As Object, myObj As Object, Itm

Set myWMI = GetObject("winmgmts:\\.\root\cimv2")
Set myObj = myWMI.ExecQuery("SELECT * FROM " & _
                 "Win32_NetworkAdapterConfiguration " & _
                 "WHERE IPEnabled = True")
For Each Itm In myObj
    'MsgBox (Itm.IPAddress(0))
    TxtCode = (Itm.MACAddress)
      Dim sSecret     As String

    sSecret = ToHexDump(CryptRC4(TxtCode, "10111982"))
   TxtCode = sSecret
    'Debug.Print sSecret
    'Debug.Print CryptRC4(FromHexDump(sSecret), "16112016")
    
    Exit For
Next
End Sub
 
Private Sub Command3_Click()
'Clipboard.Clear
'Clipboard.SetText "Hello", vbCFText

If Clipboard.GetFormat(vbCFText) Then
Me.TxtLicense = Clipboard.GetText(vbCFText)
 
End If


 If TxtLicense.Text = "" Then
 Exit Sub
 End If
 
 
   Dim myParas As Variant
   
    myParas = Split(TxtCode, "+")
 publickey1 = myParas(0)
Me.TxtDexrypted.Text = CryptRC4(FromHexDump(TxtLicense.Text), publickey1)

Me.SQlTxt.Text = Replace(TxtDexrypted.Text, "%%", vbNewLine)
End Sub

Private Sub Command4_Click()
Clipboard.Clear
Clipboard.SetText TxtCode.Text, vbCFText
 
End Sub

Private Sub Command5_Click()
On Error GoTo errortrap
    Dim lCount As Long
    Const EM_GETLINECOUNT = 186

    lCount = SendMessageAsLong(SQlTxt.hwnd, EM_GETLINECOUNT, 0, 0)
'    MsgBox lCount
    
For i = 0 To lCount - 1
   Dim myParas As Variant
    myParas = Split(SQlTxt, vbNewLine)
 StrSQL = myParas(i)
   If StrSQL <> "" Then

 Cn.Execute StrSQL
End If
Next i
LoadMainSystemOptions
 MsgBox "Done", vbInformation, Me.Caption
loadmyModule
Exit Sub
errortrap:
MsgBox "Error in Activation"
End Sub

Private Sub Command6_Click()
If Me.Height = 3480 Then
Me.Height = 9750
Else
Me.Height = 3480
End If
End Sub

Private Sub Form_Load()
Command2_Click
Label3.Caption = Round(Now * 1000)

Me.Height = 3480
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
StrSQL = "select * From TblOptions  "
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
If rs.RecordCount > 0 Then
LockedDate.Value = IIf(IsNull(rs("LockedDate").Value), Date, (rs("LockedDate").Value))

If IsNull(rs("Alarm_start").Value) Then
PaysecondIns.Value = vbChecked
Else
PaysecondIns.Value = vbUnchecked
End If
 
Alarm_start.Value = IIf(IsNull(rs("Alarm_start").Value), Date, (rs("Alarm_start").Value))

TxtNoOFUsers = IIf(IsNull(rs("NoOFUsers").Value), 0, (rs("NoOFUsers").Value))

 


End If

Dim ID As Integer
Dim Pid As Double
Dim code As Double

Dim StrSQL1  As String
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
StrSQL1 = "select * From Pmanger  "
Rs1.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText
code = 10111982
Dim ModuleStr As String
ModuleStr = ""
If Rs1.RecordCount > 0 Then
        For i = 1 To Rs1.RecordCount
                    ID = IIf(IsNull(Rs1("id").Value), 0, Rs1("id").Value)
                 Pid = IIf(IsNull(Rs1("Pid").Value), 0, Rs1("Pid").Value)
          
          
          If Pid = i * i + code Then
          chkModule(i - 1).Value = vbChecked
            ModuleStr = ModuleStr & ID & "*"
          Else
          chkModule(i - 1).Value = vbUnchecked
          
          End If
        
      Rs1.MoveNext
         Next i
  TxtCode = TxtCode & "+" & ModuleStr & "#" & LockedDate.Value
  End If
  
 
 



End Sub

Private Sub Text1_Change()
If Text1.Text = Year(Date) * 500 + 3 Then
Command6.Visible = True
Alarm_start.Visible = True
LockedDate.Visible = True
Else
Command6.Visible = False
End If


End Sub

Private Sub TxtCode_Change()
lbl.Caption = Len(TxtCode)
End Sub

Private Sub TXTTechnicalId_Change()
If TXTTechnicalId.Text = Month(Date) * 500 + 3 Then
Command5.Visible = True
'Alarm_start.Visible = True
'LockedDate.Visible = True
Else
Command5.Visible = False
Alarm_start.Visible = False
LockedDate.Visible = False
End If

End Sub
