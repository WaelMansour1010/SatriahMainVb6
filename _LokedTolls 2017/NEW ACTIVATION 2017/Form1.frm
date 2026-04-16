VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "04 02 2020"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10575
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
   Begin VB.CommandButton Command4 
      Caption         =   "Paste"
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox TxtWebAdv 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   58
      Text            =   "http://sattaryahadv.xyz/MainAdvertisement/index"
      Top             =   6120
      Width           =   5055
   End
   Begin VB.CheckBox VbEcnomy 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Lite"
      Height          =   255
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Activate"
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtLicense 
      Alignment       =   1  'Right Justify
      Height          =   1095
      Left            =   -840
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   54
      Top             =   7800
      Width           =   7935
   End
   Begin VB.TextBox TxtCode 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   53
      Top             =   1680
      Width           =   9255
   End
   Begin VB.TextBox TxtPassword 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox TxtNoOFUsers 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÇáãæÏíæáÇÊ"
      Enabled         =   0   'False
      Height          =   3735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   9615
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÍÌæÒ"
         Height          =   255
         Index           =   50
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÍÌæÒ"
         Height          =   255
         Index           =   49
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÍÌæÒ"
         Height          =   255
         Index           =   48
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÒÑÚå"
         Height          =   255
         Index           =   47
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇÓÊÞÏÇã"
         Height          =   255
         Index           =   46
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "äÞá ÇáÑßÇÈ"
         Height          =   255
         Index           =   45
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÈÕÑíÇÊ"
         Height          =   255
         Index           =   44
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÌãíá"
         Height          =   255
         Index           =   43
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÚÏÇÊ/ÇáÓíÇÑÇÊ"
         Height          =   255
         Index           =   42
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÔÇÛá"
         Height          =   255
         Index           =   41
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   " ÇáÔÆæä ÇáÞÇäæäíÉ"
         Height          =   255
         Index           =   40
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÞíãå ÇáãÖÇÝÉ"
         Height          =   255
         Index           =   39
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÚãÇÑ ÇáÏíæä"
         Height          =   255
         Index           =   38
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÄÔÑÇÊ ÇáÍíÉ"
         Height          =   255
         Index           =   37
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÍÌ æ ÇáÚãÑÉ"
         Height          =   255
         Index           =   36
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÏÇÑÉ ÇáãÕÇÚÏ"
         Height          =   255
         Index           =   35
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈíÚÇÊ ÇáÊÞÓíØ"
         Height          =   255
         Index           =   34
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÓãÇåãÇÊ ÇáÚÞÇÑíÉ"
         Height          =   255
         Index           =   33
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÔÆæä ÇáÇÏÇÑíÉ"
         Height          =   255
         Index           =   32
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÚÇãáÇÊ ÇáÈäßíÉ"
         Height          =   255
         Index           =   31
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÓæíÞ ÇáÚÞÇÑí"
         Height          =   255
         Index           =   30
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÇíÓÇÊ"
         Height          =   255
         Index           =   29
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÍÕíáÇÊ"
         Height          =   255
         Index           =   28
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkmai 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÈÑäÇãÌ ÇáÚÇã"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   29
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkselectall 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÎÊíÇÑ Çáßá"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   28
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÍÇæíÇÊ"
         Height          =   255
         Index           =   27
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇÏÇÁ æÇáãåÇã"
         Height          =   255
         Index           =   26
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇÏæÇÊ ÇáÝ"
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
         Caption         =   "ÇáÈíÇäÇÊ ÇáÇÓÇÓíÉ"
         Height          =   255
         Index           =   24
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÏíÑ ÇáäÙÇã"
         Height          =   255
         Index           =   23
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÞÇÑíÑ"
         Height          =   255
         Index           =   22
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÏÇÑÉ ÇáÇãáÇß"
         Height          =   255
         Index           =   21
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇÓåã"
         Height          =   255
         Index           =   20
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÏÇÑÓ æÇáãÚÇåÏ ÇáÊÚáíãíÉ"
         Height          =   255
         Index           =   19
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇÑÔíÝ ÇáÇáßÊÑæäí"
         Height          =   255
         Index           =   18
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáäÞá ÇáãÏÑÓí"
         Height          =   255
         Index           =   17
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÕíÇäÉ ÇáÚÇãÉ"
         Height          =   255
         Index           =   16
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÕíÇäÉ ÇáãÚÏÇÊ/ÇáÓíÇÑÇÊ"
         Height          =   255
         Index           =   15
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáäÞáíÇÊ"
         Height          =   255
         Index           =   14
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÔÇÑíÚ"
         Height          =   255
         Index           =   13
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÑÇÞÈÉ ÇáÌæÏÉ"
         Height          =   255
         Index           =   12
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇäÊÇÌ"
         Height          =   255
         Index           =   11
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "äÞÇØ ÇáÈíÚ"
         Height          =   255
         Index           =   10
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÔÍä"
         Height          =   255
         Index           =   9
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÈíÚÇÊ"
         Height          =   255
         Index           =   8
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÓæíÞ"
         Height          =   255
         Index           =   7
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÔÊÑíÇÊ"
         Height          =   255
         Index           =   6
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÎÒæä"
         Height          =   255
         Index           =   5
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "Ô ã"
         Height          =   255
         Index           =   4
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÇÕæá ÇáËÇíÊÉ"
         Height          =   255
         Index           =   3
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊÍáíá ÇáãÇáí"
         Height          =   255
         Index           =   2
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáãÚÇãáÇÊ ÇáãÇáíÉ"
         Height          =   255
         Index           =   1
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÍÓÇÈÇÊ"
         Height          =   255
         Index           =   0
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox PaysecondIns 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker LockedDate 
      Height          =   345
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   123273217
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker Alarm_start 
      Height          =   345
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   123273217
      CurrentDate     =   38784
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ ÇáÊÓÌíá"
      Height          =   255
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ ÇáÊÝÚíá"
      Height          =   255
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÚÏÏ ÇáãÓÊÎÏãíä"
      Height          =   255
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ÊÇÑíÎ ÇÓÊÍÞÇÞ ÇáÞÓØ ÇáÊÇáí"
      Height          =   255
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÎÑ ÊÇÑíÎ ÕíÇäÉ"
      Height          =   255
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PublicKey As String
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
 
Private Sub chkmai_Click(Index As Integer)
Dim i As Integer
For i = 0 To 3
 
 
 
If chkmai(Index).Value = vbChecked Then
chkModule(i).Value = vbChecked
chkModule(23).Value = vbChecked
chkModule(22).Value = vbChecked

chkModule(24).Value = vbChecked
chkModule(31).Value = vbChecked
Else
chkModule(i).Value = vbUnchecked
chkModule(23).Value = vbUnchecked
chkModule(24).Value = vbUnchecked
chkModule(31).Value = vbUnchecked
chkModule(22).Value = vbUnchecked

End If


Next i
End Sub

Private Sub chkselectall_Click(Index As Integer)
On Error Resume Next
Dim i As Integer
For i = 0 To 100
 
If chkselectall(Index).Value = vbChecked Then
chkModule(i).Value = vbChecked
Else
chkModule(i).Value = vbUnchecked
chkModule(23).Value = vbChecked
End If


Next i

End Sub

Private Sub Command1_Click()
Dim fulltext As String
If TxtPassword.Text = "" Then
Else
MsgBox "wrong code"
Exit Sub
End If
fulltext = ""
StrSQL = "update TblOptions set   Alarm_start=null,LockSystem=0, LockedDate='" & SQLDate(Me.LockedDate.Value) & "'"
'Cn.Execute StrSQL
DoEvents

fulltext = fulltext & "%%" & StrSQL
If PaysecondIns.Value = vbChecked Then
StrSQL = "update TblOptions set    Alarm_start=null"
Else
StrSQL = "update TblOptions set    Alarm_start='" & SQLDate(Me.Alarm_start.Value) & "'"
End If
DoEvents
'Cn.Execute StrSQL
fulltext = fulltext & "%%" & StrSQL

'*************
If VbEcnomy.Value = vbChecked Then
StrSQL = "update TblOptions set    Ecnomy=1"
Else
StrSQL = "update TblOptions set    Ecnomy=0"
End If
DoEvents
'Cn.Execute StrSQL
fulltext = fulltext & "%%" & StrSQL

'*************
StrSQL = "update TblOptions set    WebAdv='" & Me.TxtWebAdv.Text & "'"
 
DoEvents
'Cn.Execute StrSQL
fulltext = fulltext & "%%" & StrSQL


'************
StrSQL = "update TblOptions set   NOOFUsers=" & Val(TxtNoOFUsers.Text) & ""
'Cn.Execute StrSQL
fulltext = fulltext & "%%" & StrSQL
Dim i As Integer
StrSQL = "update Pmanger set   Pid=0"
'Cn.Execute StrSQL
fulltext = fulltext & "%%" & StrSQL
For i = 1 To 50
            If chkModule(i - 1).Value = vbChecked Then
                    StrSQL = "update Pmanger set   Pid=" & TxtPassword.Text + i * i & ""
         StrSQL = StrSQL & " where id=" & i
'         Cn.Execute StrSQL
         fulltext = fulltext & "%%" & StrSQL
            End If




Next i
If TxtCode.Text = "" Then GoTo LL
    sSecret = ToHexDump(CryptRC4(fulltext, PublicKey))
   TxtLicense = sSecret
    'Debug.Print sSecret
'TxtLicense = CryptRC4(FromHexDump(sSecret), "10111982")
    
LL:
 MsgBox "Done", vbInformation, Me.Caption

End Sub

Private Sub Command2_Click()
Command1_Click
DoEvents
Clipboard.Clear
Clipboard.SetText TxtLicense.Text, vbCFText
'If Clipboard.GetFormat(vbCFText) Then
'   Text1.Text = Clipboard.GetText(vbCFText)
'End If
End Sub

Private Sub Command3_Click()
TxtCode.Text = ""
End Sub

Private Sub Command4_Click()
Dim hashpos As Integer
On Error Resume Next
If Clipboard.GetFormat(vbCFText) Then
Me.TxtCode = Clipboard.GetText(vbCFText)
 hashpos = InStr(Replace(TxtCode.Text, Chr(13), ""), "#")
 Dim datelock As Date
 
 If hashpos > 0 Then
 
 datelock = Mid(Replace(TxtCode.Text, Chr(13), ""), hashpos + 1, 10)
 LockedDate.Value = datelock
End If
If hashpos = 0 Then Exit Sub
TxtCode.Text = Mid(TxtCode, 1, hashpos - 1)

End If


End Sub

Private Sub Form_Load()
Me.Caption = Month(Date) * 500 + 3
Dim StrSQL  As String
'Dim rs As ADODB.Recordset
'Set rs = New ADODB.Recordset
'StrSQL = "select * From TblOptions  "
'rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
'If rs.RecordCount > 0 Then
'LockedDate.Value = IIf(IsNull(rs("LockedDate").Value), Date, (rs("LockedDate").Value))

'If IsNull(rs("Alarm_start").Value) Then
'PaysecondIns.Value = vbChecked
'Else
'PaysecondIns.Value = vbUnchecked
'End If
 
'Alarm_start.Value = IIf(IsNull(rs("Alarm_start").Value), Date, (rs("Alarm_start").Value))

'TxtNoOFUsers = IIf(IsNull(rs("NoOFUsers").Value), 0, (rs("NoOFUsers").Value))

 


'End If

Dim id As Integer
Dim Pid As Double
Dim code As Double

Dim StrSQL1  As String
'Dim rs1 As ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'StrSQL1 = "select * From Pmanger  "
'rs1.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText
code = 10111982
'If rs1.RecordCount > 0 Then
'        For i = 1 To rs1.RecordCount
'                    id = IIf(IsNull(rs1("id").Value), 0, rs1("id").Value)
'                 Pid = IIf(IsNull(rs1("Pid").Value), 0, rs1("Pid").Value)
'
'
'          If Pid = i * i + code Then
'          chkModule(i - 1).Value = vbChecked
'          Else
'          chkModule(i - 1).Value = vbUnchecked
'
'          End If
          
'      rs1.MoveNext
'         Next i
'
'  End If
  
 
 


  

End Sub
Function clearAllCheck()
On Error Resume Next
        For i = 1 To 40
      
          chkModule(i - 1).Value = vbUnchecked
    
         Next i
End Function
Private Sub TxtCode_Change()
On Error Resume Next

clearAllCheck
   Dim myParas As Variant
    myParas = Split(TxtCode, "+")
 PublicKey = myParas(0)
  ModulesStr = myParas(1)
'  LockedDate.Value = Date
  Alarm_start.Value = Date
  PaysecondIns.Value = vbChecked
  
  Dim mymodule As Variant
  mymodule = Split(ModulesStr, "*")
      For i = 0 To Len(ModulesStr)
      
          chkModule(Val(mymodule(i)) - 1).Value = vbChecked
    
         Next i
         

End Sub

Private Sub TxtPassword_Change()
'10111982
If TxtPassword.Text = "Satar9090" Then
Command1.Visible = True
Command2.Visible = True

Frame1.Enabled = True
LockedDate.Enabled = True
Alarm_start.Enabled = True
TxtNoOFUsers.Enabled = True
PaysecondIns.Enabled = True
Else
Frame1.Enabled = False
Alarm_start.Enabled = False
TxtNoOFUsers.Enabled = False
PaysecondIns.Enabled = False
Command1.Visible = False
Command2.Visible = False
Exit Sub
End If

End Sub
