VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "ÇÚÏÇÏÊ ÇáÇäÙãÉ"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTDB 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Text            =   "Byte"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Backup"
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtWebAdv 
      BackColor       =   &H000000FF&
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
      Left            =   480
      TabIndex        =   58
      Text            =   "http://sattaryahadv.xyz/MainAdvertisement/index"
      Top             =   2160
      Width           =   5055
   End
   Begin VB.CheckBox VbEcnomy 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Caption         =   "Lite"
      Height          =   255
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy License"
      Height          =   495
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   6840
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
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   53
      Top             =   1680
      Width           =   5055
   End
   Begin VB.TextBox TxtPassword 
      Alignment       =   1  'Right Justify
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
      Height          =   3615
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   9615
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÍÌæÒ"
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
         Caption         =   "ãÍÌæÒ"
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
         Caption         =   "æÑÔ ÇáÏåÈ æÇáÇáãÇÓ"
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
         Caption         =   "ãÒÇÑÚ ÇáÏæÇÌä"
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
         Caption         =   "ÇáÔÞÞ ÇáÝäÏÞíÉ"
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
         Caption         =   "ÇáÊÎØíØ"
         Height          =   255
         Index           =   27
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkModule 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇáÊØæíÑ"
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
         Caption         =   "ÇáÇÑÔíÝ"
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
      Caption         =   "ÊÝÚíá"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox PaysecondIns 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10560
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Format          =   49217537
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
      Format          =   49217537
      CurrentDate     =   38784
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ ÇáÊÓÌíá"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5880
      TabIndex        =   56
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ ÇáÊÝÚíá"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6480
      TabIndex        =   38
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÚÏÏ ÇáãÓÊÎÏãíä"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ÊÇÑíÎ ÇÓÊÍÞÇÞ ÇáÞÓØ ÇáÊÇáí"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÎÑ ÊÇÑíÎ ÕíÇäÉ"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   6480
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
chkModule(24).Value = vbChecked
chkModule(31).Value = vbChecked
Else
chkModule(i).Value = vbUnchecked
chkModule(23).Value = vbUnchecked
chkModule(24).Value = vbUnchecked
chkModule(31).Value = vbUnchecked

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
If TxtPassword.Text = "10111982" Then
Else
MsgBox "wrong code"
Exit Sub
End If

 
StrSQL = "update TblOptions"
StrSQL = StrSQL & "  set Company_Name=''"
StrSQL = StrSQL & "  ,Company_Address=''"
StrSQL = StrSQL & "  ,Company_Arabic_Name=''"

Cn.Execute StrSQL
DoEvents
StrSQL = "update TblBranchesData"
StrSQL = StrSQL & "  set branch_name='1'"
StrSQL = StrSQL & "  ,branch_namee='1'"

Cn.Execute StrSQL
DoEvents

 
StrSQL = StrSQL & " Update tblActivitesType"
StrSQL = StrSQL & " set Name=''"
StrSQL = StrSQL & " ,Namee=''"


Cn.Execute StrSQL
DoEvents

  MsgBox "Done", vbInformation, Me.Caption

End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText TxtLicense.Text, vbCFText
'If Clipboard.GetFormat(vbCFText) Then
'   Text1.Text = Clipboard.GetText(vbCFText)
'End If
End Sub

Private Sub Command3_Click()
StrSQL = "BACKUP DATABASE " & Me.TXTDB & ""
StrSQL = StrSQL & " TO DISK = 'c:\" & Me.TXTDB & ".Bak'"
 
Cn.Execute StrSQL
MsgBox "Done"
End Sub

Private Sub Form_Load()
Me.Caption = Month(Date) * 500 + 3
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

Dim id As Integer
Dim Pid As Double
Dim code As Double

Dim StrSQL1  As String
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
StrSQL1 = "select * From Pmanger  "
rs1.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText
code = 10111982
If rs1.RecordCount > 0 Then
        For i = 1 To rs1.RecordCount
                    id = IIf(IsNull(rs1("id").Value), 0, rs1("id").Value)
                 Pid = IIf(IsNull(rs1("Pid").Value), 0, rs1("Pid").Value)
          
          
          If Pid = i * i + code Then
          chkModule(i - 1).Value = vbChecked
          Else
          chkModule(i - 1).Value = vbUnchecked
          
          End If
          
      rs1.MoveNext
         Next i
  
  End If
  
 
 


  

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
  
  
  Dim mymodule As Variant
  mymodule = Split(ModulesStr, "*")
      For i = 1 To Len(ModulesStr)
      
          chkModule(Val(mymodule(i)) - 1).Value = vbChecked
    
         Next i
         

End Sub

Private Sub TxtPassword_Change()
If TxtPassword.Text = "10111982" Then
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
