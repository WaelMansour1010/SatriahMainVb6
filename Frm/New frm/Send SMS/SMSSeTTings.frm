VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form SMSSeTTings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«⁄œ«œ«  —”«∆· «·ÃÊ«·"
   ClientHeight    =   8100
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13200
   Icon            =   "SMSSeTTings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   13996
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   " «Œ »«— ≈—”«· «·—”«∆·"
      TabPicture(0)   =   "SMSSeTTings.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command4"
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(2)=   "txtNumbers"
      Tab(0).Control(3)=   "txtMessage"
      Tab(0).Control(4)=   "txtSender"
      Tab(0).Control(5)=   "WbHelp"
      Tab(0).Control(6)=   "lblBalance"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Label1(0)"
      Tab(0).Control(9)=   "Label3"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "«⁄œ«œ«  «·«—”«·"
      TabPicture(1)   =   "SMSSeTTings.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label11"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label12"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtSenderName"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "CmdCheckSender"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "CmdAddSender"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "CmdRegisterSender"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "TxtActivationCode"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TxtActivecode2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "UserName"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Password"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "CMDSave"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "OPTWEB(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "OPTWEB(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "OPTWEB(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "OPTWEB(3)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "OPTWEB(4)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmdtest"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "«⁄œ«œ«  «·«Ì„Ì·« "
      TabPicture(2)   =   "SMSSeTTings.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LBLStatus"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(3)=   "cmdSend"
      Tab(2).Control(4)=   "Inet2"
      Tab(2).Control(5)=   "Command3"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdtest 
         Caption         =   "test hisms"
         Height          =   615
         Left            =   6240
         TabIndex        =   65
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   " ÕœÌÀ «·—’Ìœ"
         Height          =   855
         Left            =   -66000
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   3600
         Width           =   1455
      End
      Begin VB.OptionButton OPTWEB 
         Alignment       =   1  'Right Justify
         Caption         =   "hisms"
         Height          =   255
         Index           =   4
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   2280
         Width           =   1695
      End
      Begin VB.OptionButton OPTWEB 
         Alignment       =   1  'Right Justify
         Caption         =   "gateway.sa"
         Height          =   255
         Index           =   3
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Õ›Ÿ «·«⁄œ«œ« "
         Height          =   375
         Left            =   -65160
         TabIndex        =   60
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton OPTWEB 
         Alignment       =   1  'Right Justify
         Caption         =   "jawalbsms.ws"
         Height          =   255
         Index           =   2
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1800
         Width           =   1695
      End
      Begin InetCtlsObjects.Inet Inet2 
         Left            =   -66600
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.OptionButton OPTWEB 
         Alignment       =   1  'Right Justify
         Caption         =   "elec.sa"
         Height          =   255
         Index           =   1
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton OPTWEB 
         Alignment       =   1  'Right Justify
         Caption         =   "Mobily.ws"
         Height          =   255
         Index           =   0
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   495
         Left            =   -68160
         TabIndex        =   52
         Top             =   5520
         Width           =   1215
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
         Left            =   -74880
         TabIndex        =   42
         Top             =   960
         Width           =   11850
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
            Left            =   9885
            PasswordChar    =   "*"
            TabIndex        =   47
            Text            =   "spamkiller"
            Top             =   300
            Width           =   1800
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
            Left            =   5355
            TabIndex        =   46
            Text            =   "a.s@sattaryah.com"
            Top             =   300
            Width           =   3120
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
            TabIndex        =   45
            Text            =   "25"
            Top             =   690
            Width           =   600
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
            TabIndex        =   44
            Text            =   "mail.sattaryah.com"
            Top             =   300
            Width           =   3360
         End
         Begin VB.CheckBox chkSSL 
            Alignment       =   1  'Right Justify
            Caption         =   "Req. SSL"
            Height          =   315
            Left            =   2475
            TabIndex        =   43
            Top             =   675
            Width           =   1065
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "or 587"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   720
            Width           =   615
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
            Index           =   3
            Left            =   150
            TabIndex        =   51
            Top             =   300
            Width           =   480
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
            TabIndex        =   50
            Top             =   675
            Width           =   300
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
            Left            =   4110
            TabIndex        =   49
            Top             =   300
            Width           =   1080
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
            Left            =   9135
            TabIndex        =   48
            Top             =   300
            Width           =   690
         End
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
         Left            =   -74880
         TabIndex        =   28
         Top             =   2160
         Width           =   7890
         Begin VB.TextBox txtMsg 
            Alignment       =   1  'Right Justify
            Height          =   1335
            Left            =   1080
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Text            =   "SMSSeTTings.frx":0060
            Top             =   1560
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
            TabIndex        =   35
            Text            =   "Dynamic ERP"
            Top             =   300
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
            TabIndex        =   34
            Text            =   "ReMinder Test"
            Top             =   1075
            Width           =   6615
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
            TabIndex        =   33
            Text            =   "info@sattaryah.com"
            Top             =   225
            Width           =   2715
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
            TabIndex        =   32
            Top             =   650
            Width           =   2115
         End
         Begin VB.TextBox txtMsg1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Text            =   "SMSSeTTings.frx":0066
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
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
            TabIndex        =   30
            Text            =   "a.s@sattaryah.com"
            Top             =   700
            Width           =   2715
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   255
            Left            =   7320
            TabIndex        =   29
            Top             =   720
            Width           =   255
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
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
            TabIndex        =   41
            Top             =   225
            Width           =   765
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
            TabIndex        =   40
            Top             =   1500
            Width           =   630
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
            TabIndex        =   39
            Top             =   300
            Width           =   810
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
            TabIndex        =   38
            Top             =   1100
            Width           =   540
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
            TabIndex        =   37
            Top             =   705
            Width           =   180
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
            TabIndex        =   36
            Top             =   675
            Width           =   930
         End
      End
      Begin VB.CommandButton CMDSave 
         Caption         =   "Õ›Ÿ «·«⁄œ«œ« "
         Height          =   615
         Left            =   240
         TabIndex        =   27
         Top             =   7260
         Width           =   2055
      End
      Begin VB.TextBox Password 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   4560
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   1620
         Width           =   1575
      End
      Begin VB.TextBox UserName 
         Height          =   495
         Left            =   1560
         TabIndex        =   23
         Top             =   1620
         Width           =   1695
      End
      Begin VB.TextBox TxtActivecode2 
         Height          =   615
         Left            =   2400
         TabIndex        =   21
         Top             =   6540
         Width           =   2295
      End
      Begin VB.TextBox TxtActivationCode 
         Height          =   615
         Left            =   2400
         TabIndex        =   16
         Top             =   5820
         Width           =   2295
      End
      Begin VB.CommandButton CmdRegisterSender 
         Caption         =   " √ﬂÌœ ÕÃ“ «”„ «·„—”·"
         Height          =   615
         Left            =   2400
         TabIndex        =   15
         Top             =   7260
         Width           =   2535
      End
      Begin VB.CommandButton CmdAddSender 
         Caption         =   "ÕÃ“ «”„ «·„—”· ⁄·Ï «·„Êﬁ⁄"
         Height          =   615
         Left            =   2760
         TabIndex        =   14
         Top             =   4380
         Width           =   2535
      End
      Begin VB.CommandButton CmdCheckSender 
         Caption         =   "«· «ﬂœ «‰ «”„ «·„—”· €Ì— „ÕÃÊ“"
         Height          =   615
         Left            =   2760
         TabIndex        =   13
         Top             =   2940
         Width           =   2535
      End
      Begin VB.TextBox TxtSenderName 
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Text            =   "mobily.ws"
         Top             =   1020
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send SMS"
         Height          =   495
         Left            =   -72840
         TabIndex        =   6
         Top             =   7380
         Width           =   2535
      End
      Begin VB.TextBox txtNumbers 
         Height          =   2415
         Left            =   -73320
         TabIndex        =   5
         Text            =   "966541793243"
         Top             =   4740
         Width           =   4455
      End
      Begin VB.TextBox txtMessage 
         Height          =   2415
         Left            =   -73320
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2100
         Width           =   4455
      End
      Begin VB.TextBox txtSender 
         Height          =   495
         Left            =   -73320
         TabIndex        =   3
         Text            =   "mobily.ws"
         Top             =   1380
         Width           =   4455
      End
      Begin XtremeSuiteControls.WebBrowser WbHelp 
         Height          =   9675
         Left            =   -61200
         TabIndex        =   57
         Top             =   960
         Visible         =   0   'False
         Width           =   2925
         _Version        =   786432
         _ExtentX        =   5159
         _ExtentY        =   17066
         _StockProps     =   173
         BackColor       =   -2147483643
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Êﬁ⁄"
         Height          =   255
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LBLStatus 
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
         Left            =   -74760
         TabIndex        =   53
         Top             =   5520
         Width           =   6390
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PassWord"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3480
         TabIndex        =   25
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UserName"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬂÊœ «· ›⁄Ì· «·„—”· ⁄·Ï «·ÃÊ«·"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5400
         TabIndex        =   22
         Top             =   6540
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "»⁄œ ÕÃ“ «”„ «·„—”· ⁄·Ï «·„Êﬁ⁄ Ì „ «—Ã«⁄ ﬂÊœ «· ›⁄Ì· «·–Ì ”Ì „ «—”«·Â „⁄ «”„ «·„—”· · √ﬂÌœ ÕÃ“ «·«”„"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   960
         TabIndex        =   20
         Top             =   5220
         Width           =   5415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "›Ì Õ«·… ⁄œ„ ÊÃÊœ «”„ «·„—”· Ì „ ÕÃ“ «”„ «·„—”· ⁄·Ï «·„Êﬁ⁄"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   960
         TabIndex        =   19
         Top             =   3780
         Width           =   5415
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ÌÃ» «Ê·« «· √ﬂœ „‰ ⁄œ„ ÊÃÊœ «”„ «·„—”· Ê«‰Â €Ì— „ÕÃÊ“ „”»ﬁ«"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1560
         TabIndex        =   18
         Top             =   2340
         Width           =   4575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "—ﬁ„  ”ÃÌ· «”„ «·„—”·"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5400
         TabIndex        =   17
         Top             =   5820
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sender"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblBalance 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -72720
         TabIndex        =   10
         Top             =   780
         Width           =   2775
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Message"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74640
         TabIndex        =   9
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sender"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   -74640
         TabIndex        =   8
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numbers"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74640
         TabIndex        =   7
         Top             =   4740
         Width           =   1215
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1440
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "SMSSeTTings.frx":006D
      Left            =   2640
      List            =   "SMSSeTTings.frx":0227
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "SMSSeTTings.frx":03E3
      Left            =   4560
      List            =   "SMSSeTTings.frx":059D
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "SMSSeTTings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''#################'''''''''''''''''''''''
'''''''''''''''#               #'''''''''''''''''''''''
'''''''''''''''# www.mobily.ws #'''''''''''''''''''''''
'''''''''''''''#               #'''''''''''''''''''''''
'''''''''''''''#################'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================
'Const UserName = "966541793243" ' Enter Your User Name Here
'Const Password = "Spamkiller16112016" ' Enter Your Password Here
'===================================================================
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const INTERNET_MAX_URL_LENGTH As Long = 2048
Private Const URL_ESCAPE_PERCENT As Long = &H1000&

Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" ( _
    ByVal pszUrl As String, _
    ByVal pszEscaped As String, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszUrl As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long

Dim rs As ADODB.Recordset

Private Sub CmdAddSender_Click()
    On Error Resume Next
    Dim s As String
    Dim Result As String

    's = "http://www.mobily.ws/api/addSender.php?mobile=" & UserName & "&password=" & Password & "&sender=" & TxtSenderName.Text
s = "http://alfa-cell.com/api/addSender.php?mobile=" & txtUsername.text & "&password=" & txtPassword.text & "&sender=" & TxtSenderName.text

    Result = Inet1.OpenURL(s)

    Select Case Result

        Case "1": MsgBox ("≈‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ··œŒÊ· ≈·Ï Õ”«» «·—”«∆· €Ì— ’ÕÌÕ ( √ﬂœ „‰ √‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ÂÊ ‰›”Â «·–Ì  ” Œœ„Â ⁄‰œ œŒÊ·ﬂ ≈·Ï „Êﬁ⁄ „Ê»«Ì·Ì)") 'mobile Not found

        Case "2": MsgBox ("Œÿ√ ›Ì ﬂ·„… «·„—Ê— ( √ﬂœ „‰ √‰ ﬂ·„… «·„—Ê— «· Ì  „ ≈” Œœ«„Â« ÂÌ ‰›”Â« «· Ì  ” Œœ„Â« ⁄‰œ œŒÊ·ﬂ „Êﬁ⁄ „Ê»«Ì·Ì,≈–« ‰”Ì  ﬂ·„… «·„—Ê— ≈÷€ÿ ⁄·Ï —«»ÿ ‰”Ì  ﬂ·„… «·„—Ê— · ’·ﬂ —”«·… ⁄·Ï ÃÊ«·ﬂ »—ﬁ„ «·„—Ê— «·Œ«’ »ﬂ)") 'error password

        Case "3": MsgBox ("≈‰ —ﬁ„ «·ÃÊ«· «·–Ì  „ ≈œŒ«·Â ·ÌﬂÊ‰ ≈”„ „—”· ·Ì” ’ÕÌÕ«(Ì—ÃÏ «· √ﬂœ „‰ ’Õ… «·—ﬁ„ Ê√‰Â »«·’Ì€… «·œÊ·Ì… „À«· (966500000000)") 'not international  mobile sender

        Case "4": MsgBox ("≈”„ «·„—”· «·–Ì √œŒ· Â ·Ì” »Õ«Ã… · ›⁄Ì·") 'not need to active

        Case "5": MsgBox ("—’Ìœﬂ €Ì— ﬂ«›Ì ·√ „«„ «·⁄„·Ì… (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance  not  enough

        Case "-1": MsgBox "·„ Ì „ «· Ê«’· „⁄ Œ«œ„ (Server) «·≈—”«· „Ê»«Ì·Ì »‰Ã«Õ. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"

        Case "-2": MsgBox "·„ Ì „ «·—»ÿ „⁄ ﬁ«⁄œ… «·»Ì«‰«  (Database) «· Ì  Õ ÊÌ ⁄·Ï Õ”«»ﬂ Ê»Ì«‰« ﬂ ·œÏ „Ê»«Ì·Ì. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
    End Select

    TxtActivationCode.text = Split(Result, "#")(1)

End Sub

Private Sub CmdCheckSender_Click()
    On Error Resume Next
    Dim s      As String
    Dim Result As String
    If OPTWEB(4) Then
       MsgBox "Â–« «·„Êﬁ⁄ ·« Ìœ⁄„ Â–Â «·Œ«’ÌÂ"
    Else

        's = "http://www.mobily.ws/api/checkSender.php?mobile=" & UserName & "&password=" & Password & "&senderId=" & TxtSenderName.Text
        s = "http://alfa-cell.com/api/checkSender.php?mobile=" & txtUsername.text & "&password=" & txtPassword.text & "&senderId=" & TxtSenderName.text
    
        Result = Inet1.OpenURL(s)

        Select Case Result

            Case "0", "": MsgBox ("«”„ «·„—”· €Ì— „›⁄·") 'new sender

            Case "1": MsgBox ("«”„ «·„—”· „” Œœ„") 'accepted sender

            Case "2": MsgBox ("«”„ «·„—”· „—›Ê÷") 'rejected  sender

            Case "3": MsgBox ("«”„ «·„” Œœ„ €Ì— „ÊÃÊœ") 'mobile Not found

            Case "4": MsgBox ("ﬂ·„… «·„—Ê— Œÿ√") 'error password

            Case Else: MsgBox (Result)
        End Select
    End If

End Sub

Private Sub CmdRegisterSender_Click()
    On Error Resume Next
    Dim s As String
    Dim Result As String

'    s = "http://www.mobily.ws/api/activeSender.php?mobile=" & UserName & "&password=" & Password & "&senderId=" & TxtActivationCode.Text & "&activeKey=" & TxtActivecode2.Text
s = "http://alfa-cell.com/api/activeSender.php?mobile=" & txtUsername.text & "&password=" & txtPassword.text & "&senderId=" & TxtActivationCode.text & "&activeKey=" & TxtActivecode2.text

    Result = Inet1.OpenURL(s)

    Select Case Result

        Case "1": MsgBox ("≈‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ··œŒÊ· €Ì— ’ÕÌÕ ( √ﬂœ „‰ √‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ÂÊ ‰›”Â «·–Ì  œŒ·Â ⁄‰œ œŒÊ·ﬂ ≈·Ï „Êﬁ⁄ „Ê»«Ì·Ì)") 'mobile Not found

        Case "2": MsgBox ("Â‰«ﬂ Œÿ√ ›Ì ﬂ·„… «·„—Ê— ( √ﬂœ „‰ √‰ ﬂ·„… «·„—Ê— «· Ì  ” Œœ„Â« ÂÌ ‰›”Â« «· Ì  ” Œœ„Â« ⁄‰œ œŒÊ·ﬂ „Êﬁ⁄ „Ê»«Ì·Ì)") 'error password

        Case "3": MsgBox (" „  ›⁄Ì· ≈”„ «·„—”· »‰Ã«Õ") 'activated  sender

        Case "4": MsgBox ("Â‰«ﬂ Œÿ√ ›Ì ﬂÊœ «· ›⁄Ì· «·–Ì  „ ≈—”«·Â. (⁄·Ìﬂ «· √ﬂœ „‰ √‰ ﬂÊœ «· ›⁄Ì· ’ÕÌÕ √Ê „—«Ã⁄… «·œ⁄„ «·›‰Ì ·≈⁄«œ… ≈—”«· ﬂÊœ «· ›⁄Ì· „—… √Œ—Ï)") 'error in activeKey

        Case "5": MsgBox ("Â‰«ﬂ Œÿ√ ›Ì —ﬁ„ ≈”„ «·„—”· «·–Ì  „ ≈—”«·Â") 'error in sender id

        Case "-1": MsgBox "·„ Ì „ «· Ê«’· „⁄ Œ«œ„ (Server) «·≈—”«· „Ê»«Ì·Ì »‰Ã«Õ. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"

        Case "-2": MsgBox "·„ Ì „ «·—»ÿ „⁄ ﬁ«⁄œ… «·»Ì«‰«  (Database) «· Ì  Õ ÊÌ ⁄·Ï Õ”«»ﬂ Ê»Ì«‰« ﬂ ·œÏ „Ê»«Ì·Ì. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
    End Select

End Sub

Private Sub CmdSave_Click()
    Dim Msg As String
    '    On Error GoTo ErrTrap

    rs("SMSUserName").value = UserName.text
    rs("SMSPassWord").value = Password.text
    rs("SenderName").value = TxtSenderName.text
    If OPTWEB(0).value = True Then
        rs("OPTWEB").value = 0
    ElseIf OPTWEB(1).value = True Then
        rs("OPTWEB").value = 1
    ElseIf OPTWEB(2).value = True Then
        rs("OPTWEB").value = 2
    ElseIf OPTWEB(3).value = True Then
        rs("OPTWEB").value = 3
    ElseIf OPTWEB(4).value = True Then
        rs("OPTWEB").value = 4
    Else
        rs("OPTWEB").value = 0
    End If
    rs.update

    LoadMainSystemOptions
    '   Unload Me
    Exit Sub
ErrTrap:
    Msg = "ÕœÀ  „‘ﬂ·… ≈À‰«¡ Õ›Ÿ «·≈⁄œ«œ« ...!!!!"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Sub

Private Sub cmdSend_Click()
 Dim RetVal          As String
    Dim objControl      As Control
    'Validate first
    For Each objControl In Me.Controls
        If TypeOf objControl Is TextBox Then
            If Trim$(objControl.text) = vbNullString And LCase$(objControl.Name) <> "txtattach" Then
              '  LBLStatus.Caption = "Error: All fields are required!"
            '    Exit Sub
            End If
        End If
    Next
    
    'Send
    Frame1.Enabled = False
    Frame2.Enabled = False
    cmdSend.Enabled = False
    LBLStatus.Caption = "Sending..."
    RetVal = SendMail(Trim$(txtTo.text), _
        Trim$(txtSubject.text), _
        Trim$(txtFromName.text) & "<" & Trim$(txtFromEmail.text) & ">", _
        Trim$(txtMsg.text), _
        Trim$(txtServer.text), _
        CInt(Trim$(txtPort.text)), _
        Trim$(txtUsername.text), _
        Trim$(txtPassword.text), _
        Trim$(txtAttach.text), _
        CBool(chkSSL.value))
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdSend.Enabled = True
    LBLStatus.Caption = IIf(RetVal = "ok", "Message sent!", RetVal)
    



End Sub

Private Sub cmdtest_Click()
    s = WebUrl4_hisms & "?get_balance&username=" & UserName.text & "&password=" & Password.text & ""
    Dim Req
    Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")
    Req.Open "get", s, async:=False
    Req.send
    Result = Req.responseText
    If InStr(1, Result, "-", vbTextCompare) Then
        Result = Split(Result, "-")(0)
    End If

    Select Case Result
         
        Case 1
            MsgBox "«”„ «·„” Œœ„ €Ì— ’ÕÌÕ "
        Case 2
            MsgBox "ﬂ·„Â «·„—Ê— €Ì— ’ÕÌÕÂ "
        Case 404
            MsgBox " „  Ã«Ê“ ⁄œœ «·„Õ«Ê·«  «·„”„ÊÕÂ "
        Case 504
            MsgBox "«·Õ”«» „⁄ÿ· "
        Case Else
            MsgBox Req.responseText
    End Select
End Sub

Private Sub Command1_Click()
    If OPTWEB(4) Then
        s = WebUrl4_hisms & "?send_sms&username=" & UserName & "&password=" _
           & Password & "&numbers=" & txtNumbers & "&sender=" & txtSender & "&message=" & txtMessage
        Dim Req
        Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")
        Req.Open "get", s, async:=False
        Req.send
        Result = Req.responseText
        If InStr(1, Result, "-", vbTextCompare) Then
            Result = Split(Result, "-")(0)
        End If

        Select Case Result
            Case "1"
                MsgBox "«”„ «·„” Œœ„ €Ì— ’ÕÌÕ"
            Case "2"
                MsgBox "ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ… "
            Case "404"
                MsgBox "·„ Ì „ «œŒ«· Ã„Ì⁄ «·»—„ —«  «·„ÿ·Ê»… "
            Case "504"
                MsgBox "«·Õ”«» „⁄ÿ· "
            Case "3"
                MsgBox " „ «·«—”«· "
            Case "4"
                MsgBox " ·«  ÊÃœ «—ﬁ«„ "
            Case "5"
                MsgBox "·«  ÊÃœ —”«·Â"
            Case "6"
                MsgBox "„—”· Œ«ÿÏ¡ "
            Case "7"
                MsgBox " „—”· €Ì— „›⁄· "
            Case "8"
                MsgBox " «·—”«·Â »Â« ﬂ·„«  „„‰Ê⁄Â "
            Case "9"
                MsgBox " ·« ÌÊÃœ —’Ìœ "
            Case "10"
                MsgBox " «—ÌŒ Œ«ÿÏ¡ "
            Case "11"
                MsgBox "Êﬁ  Œ«ÿÏ¡ "

            Case Else
                MsgBox Req.responseText
        End Select
     
    Else
        SendMessage
        updateBalance
    End If
  
End Sub

Public Sub SendMessage(Optional msgstr As String = "", _
                       Optional Numbers As String = "")
    Dim t As String

    If msgstr = "" Then
        msgstr = txtMessage.text
    End If

    If Numbers = "" Then
        Numbers = txtNumbers.text
    End If

    ''t = send(UserName, URLEncode(Password), ConvertToUnicode(ConvertString(txtMessage.Text)), txtSender.Text, txtNumbers.Text)
    If OPTWEB(0).value = True Then
 
        t = send(UserName, URLEncode(Password), ConvertToUnicode(msgstr), txtSender.text, Numbers)
    ElseIf OPTWEB(1).value = True Then
        t = send(UserName, (Password), (msgstr), txtSender.text, Numbers)
 
    ElseIf OPTWEB(2).value = True Then
        t = send(UserName, (Password), (msgstr), txtSender.text, Numbers)
 
    ElseIf OPTWEB(3).value = True Then
        t = send(UserName, (Password), (msgstr), txtSender.text, Numbers)
 
    ElseIf OPTWEB(4).value = True Then
  
       ' t = sendMessageM(UserName, Password, URLEncode(msgstr), txtSender.text, Numbers)
       
 
    End If
 
    If msgstr = "" Then
        ShowResult (t)
    Else
        ShowResult t, 1
    End If

End Sub

Public Function send(UserName As String, _
                     Password As String, _
                     Msg As String, _
                     sender As String, _
                     Numbers As String) As String
    On Error Resume Next
    Dim s As String
 
    UserName = SystemOptions.SMSUserName
    Password = (SystemOptions.SMSPassWord)
    sender = SystemOptions.SenderName
    If OPTWEB(0).value = True Then
        's = "http://www.mobily.ws/api/msgSend.php?mobile=" & UserName & "&password=" & Password & "&numbers=" & Numbers & "&sender=" & sender & "&msg=" & Msg & "&applicationType=24"
        s = "http://alfa-cell.com/api/msgSend.php?mobile=" & UserName & "&password=" & Password & "&numbers=" & Numbers & "&sender=" & sender & "&msg=" & Msg & "&applicationType=72"

        send = Inet1.OpenURL(s)
    ElseIf OPTWEB(1).value = True Then
        s = "http://elec.sa/sms/api/sendsms.php?username=" & UserName & "&password=" & Password & "&message=" & Msg & "&numbers=" & Numbers & "&sender=" & sender & "&unicode=UTF-8&return=string"
 
        send = WebRequest(s)

        MsgBox send

    ElseIf OPTWEB(2).value = True Then
        's = "http://elec.sa/sms/api/sendsms.php?username=" & UserName & "&password=" & Password & "&message=" & Msg & "&numbers=" & Numbers & "&sender=" & sender & "&unicode=UTF-8&return=string"
        s = "http://www.jawalbsms.ws/api.php/sendsms?user=" & UserName & "&pass=" & Password & "&to=" & Numbers & "&message= " & Msg & " &sender=" & sender & ""
        send = WebRequest(s)

        MsgBox send

    ElseIf OPTWEB(3).value = True Then
 
        's = "http://www.jawalbsms.ws/api.php/sendsms?user=" & UserName & "&pass=" & Password & "&to=" & Numbers & "&message= " & Msg & " &sender=" & sender & ""
        s = "https://apps.gateway.sa/vendorsms/pushsms.aspx?user=" & UserName & "&password=" & Password & "&msisdn=" & Numbers & "&sid=" & sender & "&msg" & Msg & "&fl=0"

        send = WebRequest(s)
        MsgBox send

    ElseIf OPTWEB(4).value = True Then
        s = "https://www.hisms.ws/api.php?send_sms&username=" & UserName & "&password=" & Password & "&numbers=" & Numbers & "&sender=" & sender & "&message=" & Msg '& "&date=2015-1-30&time=24:01"
        send = WebRequest(s)

        MsgBox send

    End If

    '”ÌÌÌÌÌÌÌÌ
End Function

Function GetBalance(UserName As String, Password As String) As String
  '  On Error Resume Next
    Dim s As String
    
    If OPTWEB(0).value = True Then
    s = "http://www.mobily.ws/api/balance.php?mobile=" & UserName & "&password=" & Password
    GetBalance = Inet1.OpenURL(s)
    ElseIf OPTWEB(1).value = True Then
    
    s = "http://elec.sa/sms/api/getbalance.php?username=" & UserName & "&password=" & Password & "&hangedBalance=true"
GetBalance = WebRequest(s)
    ElseIf OPTWEB(4).value = True Then
    
   ' s = "http://elec.sa/sms/api/getbalance.php?username=" & UserName & "&password=" & Password & "&hangedBalance=true"
    s = "http://www.hisms.ws/api.php?get_balance&username=" & UserName & "&password=" & Password & ""
GetBalance = WebRequestPHP(s, True)

    End If
    
End Function

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

Private Function isArabic(val As String) As Boolean

    Dim i As Integer
    Dim str As String
    str = "œÃÕŒÂ⁄€›ﬁÀ’÷ÿﬂ„‰ «·»Ì”‘Ÿ“Ê…Ï·«—ƒ¡∆≈·≈√·√¬·¬"

    For i = 0 To Len(val)

        If InStr(0, str, mId(val, i, 1), vbTextCompare) <> 0 Then
            isArabic = True
        End If

    Next i

    isArabic = False
           
End Function

Function ConvertString(s As String) As String
    Dim Arr() As String
    Dim i As Integer
    Arr = Split(s, vbNewLine)
    Dim st As String

    For i = 0 To UBound(Arr)
        st = st & Arr(i) & "'"
    Next

    ConvertString = st
End Function


Function toUnicode(Ch As String) As String
    Dim i As Integer

    If Ch = "'" Then
        toUnicode = "000D"
        Exit Function
    End If

    For i = 0 To List1.ListCount - 1

        If Ch = List1.List(i) Then
            toUnicode = List2.List(i)
            Exit Function
        End If

    Next

End Function

Private Sub Command2_Click()
CD1.ShowOpen
txtAttach.text = CD1.FileName
End Sub

Private Sub Command3_Click()
    rs("cdoSMTPServer").value = txtServer.text
    rs("cdoSMTPServerPort").value = txtPort.text
            If chkSSL.value = True Then
                    rs("cdoSMTPUseSSL").value = 1
         Else
                    rs("cdoSMTPUseSSL").value = 0
        End If

     rs("cdoSendUserName").value = txtUsername.text
     
     rs("cdoSendPassword").value = txtPassword.text
     rs("txtFromName").value = txtFromName.text
     rs("txtFromEmail").value = txtFromEmail.text
     
      
    rs.update
    MsgBox " „ «·Õ›Ÿ", vbInformation
End Sub

Private Sub Command4_Click()
    If OPTWEB(4) Then
        s = WebUrl4_hisms & "?get_balance&username=" & UserName.text & "&password=" & Password.text & ""
        Dim Req
        Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")
        Req.Open "get", s, async:=False
        Req.send
        Result = Req.responseText
        If InStr(1, Result, "-", vbTextCompare) Then
            Result = Split(Result, "-")(0)
        End If
 
        Select Case Result
         
            Case 1
                MsgBox "«”„ «·„” Œœ„ €Ì— ’ÕÌÕ "
            Case 2
                MsgBox "ﬂ·„Â «·„—Ê— €Ì— ’ÕÌÕÂ "
            Case 404
                MsgBox " „  Ã«Ê“ ⁄œœ «·„Õ«Ê·«  «·„”„ÊÕÂ "
            Case 504
                MsgBox "«·Õ”«» „⁄ÿ· "
            Case Else
                MsgBox Req.responseText
        End Select
    Else
        updateBalance
    End If

End Sub

Private Sub Form_Load()
   ' On Error GoTo ErrTrap
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Dim i As Integer

    List1.AddItem "'"
    List1.AddItem " "

    For i = 0 To List2.ListCount - 1
        List2.List(i) = Fourdigit(List2.List(i))
    Next

    List2.AddItem "000D"
    List2.AddItem "0020"

    Set rs = New ADODB.Recordset
    rs.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rs.EOF Or rs.BOF) Then
    
    txtServer.text = IIf(IsNull(rs("cdoSMTPServer").value), "", rs("cdoSMTPServer").value)
    txtPort.text = IIf(IsNull(rs("cdoSMTPServerPort").value), 587, rs("cdoSMTPServerPort").value)
    
    txtUsername.text = IIf(IsNull(rs("cdoSendUserName").value), "", rs("cdoSendUserName").value)
    txtPassword.text = IIf(IsNull(rs("cdoSendPassword").value), "", rs("cdoSendPassword").value)
    txtFromName.text = IIf(IsNull(rs("txtFromName").value), "", rs("txtFromName").value)
    txtFromEmail.text = IIf(IsNull(rs("txtFromEmail").value), "", rs("txtFromEmail").value)
      
      If IsNull(rs("cdoSMTPUseSSL").value) Then
      chkSSL.value = False
      Else
                    If rs("cdoSMTPUseSSL").value = 0 Then
                    chkSSL.value = False
                    Else
                    chkSSL.value = True
                    End If
      
      End If
    
     
      
    '************************sms******************
     
     
   '  rs("SMSUserName").value = txtUsername.Text
   '
   '  rs("SMSPassWord").value = txtPassword.Text
   '  rs("txtFromName").value = TxtSenderName.Text
   '  rs("txtFromEmail").value = txtFromEmail.Text
     
     
     
     
     
        TxtSenderName.text = IIf(IsNull(rs("SenderName").value), "", rs("SenderName").value)
        txtSender.text = IIf(IsNull(rs("SenderName").value), "", rs("SenderName").value)

        UserName.text = IIf(IsNull(rs("SMSUserName").value), "", rs("SMSUserName").value)
        Password.text = IIf(IsNull(rs("SMSPassWord").value), "", rs("SMSPassWord").value)
  
   If IsNull(rs("OPTWEB").value) Then
   OPTWEB(0).value = True
   Else
                If rs("OPTWEB").value = 0 Then
                OPTWEB(0).value = True
                ElseIf rs("OPTWEB").value = 1 Then
                OPTWEB(1).value = True
                 
               ElseIf rs("OPTWEB").value = 2 Then
                OPTWEB(2).value = True
                
                
               ElseIf rs("OPTWEB").value = 3 Then
                OPTWEB(3).value = True
                          ElseIf rs("OPTWEB").value = 4 Then
                OPTWEB(4).value = True
                End If
                
   
   End If
  
    End If

    updateBalance
    Exit Sub
ErrTrap:
End Sub

Private Sub updateBalance()
If UserName.text = "" Then Exit Sub
    Dim b As String
    If UserName = "" Then lblBalance.Caption = 0: Exit Sub
    b = GetBalance(UserName, Password)
     lblBalance.Caption = "  " & b
If UserName = "" Then lblBalance.Caption = 0: Exit Sub
    Select Case b

        Case "1"
            MsgBox "≈‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ··œŒÊ· €Ì— ’ÕÌÕ ( √ﬂœ „‰ √‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ÂÊ ‰›”Â «·–Ì  œŒ·Â ⁄‰œ œŒÊ·ﬂ ≈·Ï „Êﬁ⁄ „Ê»«Ì·Ì)" 'Mobile not found
            Exit Sub

        Case "2"
            MsgBox "Â‰«ﬂ Œÿ√ ›Ì ﬂ·„… «·„—Ê— ( √ﬂœ „‰ √‰ ﬂ·„… «·„—Ê— «· Ì  ” Œœ„Â« ÂÌ ‰›”Â« «· Ì  ” Œœ„Â« ⁄‰œ œŒÊ·ﬂ „Êﬁ⁄ „Ê»«Ì·Ì)" 'password error
            Exit Sub

        Case "-1"
            MsgBox "·„ Ì „ «· Ê«’· „⁄ Œ«œ„ (Server) «·≈—”«· „Ê»«Ì·Ì »‰Ã«Õ. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
            Exit Sub

        Case "-2"
            MsgBox "·„ Ì „ «·—»ÿ „⁄ ﬁ«⁄œ… «·»Ì«‰«  (Database) «· Ì  Õ ÊÌ ⁄·Ï Õ”«»ﬂ Ê»Ì«‰« ﬂ ·œÏ „Ê»«Ì·Ì. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
            Exit Sub
    End Select

    lblBalance.Caption = "Your Balance is: " & b
End Sub

Function Fourdigit(Ch As String) As String

    If Len(Ch) = 1 Then
        Fourdigit = "000" & Ch
        Exit Function
    End If

    If Len(Ch) = 2 Then
        Fourdigit = "00" & Ch
        Exit Function
    End If

    If Len(Ch) = 3 Then
        Fourdigit = "0" & Ch
        Exit Function
    End If

    If Len(Ch) = 4 Then
        Fourdigit = Ch
        Exit Function
    End If

End Function
Public Function URLEncode( _
    ByVal URL As String, _
    Optional ByVal SpacePlus As Boolean = True) As String
    
    Dim cchEscaped As Long
    Dim HRESULT As Long
    
    If Len(URL) > INTERNET_MAX_URL_LENGTH Then
        Err.Raise &H8004D700, "URLUtility.URLEncode", _
                  "URL parameter too long"
    End If
    
    cchEscaped = Len(URL) * 1.5
    URLEncode = String$(cchEscaped, 0)
    HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT = E_POINTER Then
        URLEncode = String$(cchEscaped, 0)
        HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    End If

    If HRESULT <> S_OK Then
'        Err.Raise Err.LastDllError, "URLUtility.URLEncode", _
'                  "System error"
    End If
    
    URLEncode = left$(URLEncode, cchEscaped)
    If SpacePlus Then
        URLEncode = Replace$(URLEncode, "+", "%2B")
        URLEncode = Replace$(URLEncode, " ", "+")
    End If
End Function


Function URLEncode2(ByVal str As String) As String
    Dim intLen As Integer
    Dim X As Integer
    Dim curChar As Long
    Dim newStr As String

    intLen = Len(str)
    newStr = ""

    For X = 1 To intLen
        curChar = Asc(mId$(str, X, 1))
          
        If (curChar < 48 Or curChar > 57) And (curChar < 65 Or curChar > 90) And (curChar < 97 Or curChar > 122) Then
            newStr = newStr & "%" & Hex(curChar)
        Else
            newStr = newStr & CHR(curChar)
        End If

    Next X
              
    URLEncode2 = newStr
End Function

