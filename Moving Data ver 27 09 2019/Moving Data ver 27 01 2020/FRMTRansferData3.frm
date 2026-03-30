VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMTRansferData3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Offline"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13590
   Icon            =   "FRMTRansferData3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14_Debug 
      Caption         =   "Debug"
      Height          =   375
      Left            =   4080
      TabIndex        =   98
      Top             =   7290
      Width           =   3555
   End
   Begin VB.CommandButton cmdTestUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3990
      TabIndex        =   97
      Top             =   8280
      Width           =   3555
   End
   Begin VB.CommandButton cmdTestInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   4080
      TabIndex        =   96
      Top             =   7740
      Width           =   3555
   End
   Begin VB.CommandButton cmdTransferMove 
      Caption         =   "«· ÕÊÌ·«  «·„Œ“‰Ì… ›ﬁÿ"
      Height          =   525
      Left            =   330
      TabIndex        =   95
      Top             =   8370
      Width           =   3615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "‰ﬁ· «·›Ê« Ì— ›ﬁÿ"
      Height          =   375
      Left            =   390
      TabIndex        =   93
      Top             =   7020
      Width           =   3555
   End
   Begin VB.CommandButton CommandTest 
      Caption         =   "Command14"
      Height          =   285
      Left            =   7440
      TabIndex        =   92
      Top             =   6240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   6330
      Top             =   5010
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "‰ﬁ· «·»Ì«‰« "
      Height          =   525
      Left            =   330
      TabIndex        =   90
      Top             =   7440
      Width           =   3615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "‰ﬁ· «·»Ì«‰« "
      Height          =   375
      Left            =   4260
      RightToLeft     =   -1  'True
      TabIndex        =   89
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   405
      Left            =   6120
      TabIndex        =   88
      Top             =   6270
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Command15 
      Caption         =   "‰ﬁ· «·»Ì«‰« "
      Height          =   375
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   6300
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command11"
      Height          =   405
      Left            =   6360
      TabIndex        =   86
      Top             =   6840
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   255
      Left            =   270
      TabIndex        =   85
      Top             =   6930
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   465
      Left            =   210
      TabIndex        =   84
      Top             =   8160
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2970
      TabIndex        =   83
      Text            =   "Text6"
      Top             =   6780
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.CommandButton Command8 
      Caption         =   " ÕœÌÀ «·⁄„·«¡ „‰ «·„⁄—÷ «·Ï «·„’‰⁄"
      Height          =   495
      Left            =   120
      TabIndex        =   81
      Top             =   6360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtCountSalesOfeers 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5250
      TabIndex        =   79
      Top             =   4770
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CheckBox chkDontCopyIss 
      Caption         =   "⁄œ„ ‰ﬁ· ”‰œ«  «·’—›"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   77
      Top             =   120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4440
      TabIndex        =   76
      Text            =   "Text5"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   " ÕœÌÀ Ãœ«Ê· «·Õ—ﬂ«  «·«”«”Ì…"
      Height          =   465
      Left            =   5580
      TabIndex        =   73
      Top             =   5760
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   9150
      TabIndex        =   70
      Top             =   9000
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   2325
      Left            =   7350
      TabIndex        =   59
      Top             =   6600
      Visible         =   0   'False
      Width           =   6045
      Begin VB.CommandButton Command6 
         Caption         =   "÷»ÿ «·”Ì—›—"
         Height          =   495
         Left            =   3450
         TabIndex        =   68
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Caption         =   "11"
         Height          =   435
         Left            =   150
         TabIndex        =   71
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton cmdUpdateSerial 
         Caption         =   "÷»ÿ «·”Ì—Ì«·"
         Height          =   495
         Left            =   3450
         TabIndex        =   69
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox txtYear 
         Height          =   345
         Left            =   2820
         TabIndex        =   66
         Text            =   "2020"
         Top             =   870
         Width           =   1035
      End
      Begin VB.TextBox txtMonth 
         Height          =   345
         Left            =   3870
         TabIndex        =   64
         Text            =   "1"
         Top             =   870
         Width           =   1035
      End
      Begin VB.TextBox txtBranch 
         Height          =   285
         Left            =   2220
         TabIndex        =   62
         Text            =   "1"
         Top             =   330
         Width           =   1065
      End
      Begin VB.OptionButton optPurs 
         Caption         =   "„‘ —Ì« "
         Height          =   195
         Left            =   3480
         TabIndex        =   61
         Top             =   150
         Width           =   1095
      End
      Begin VB.OptionButton optSales 
         Caption         =   "„»Ì⁄« "
         Height          =   195
         Left            =   4860
         TabIndex        =   60
         Top             =   150
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DCboBranch 
         Height          =   315
         Left            =   90
         TabIndex        =   72
         Top             =   330
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label14 
         Caption         =   "«·”‰…"
         Height          =   225
         Left            =   3240
         TabIndex        =   67
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "«·‘Â—"
         Height          =   225
         Left            =   4080
         TabIndex        =   65
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "«·›—⁄"
         Height          =   225
         Left            =   2370
         TabIndex        =   63
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdUdateFiles 
      Caption         =   " ÕœÌÀ «·„·›«  «·«”«”Ì…"
      Height          =   495
      Left            =   330
      TabIndex        =   58
      Top             =   7920
      Width           =   3615
   End
   Begin VB.CommandButton cmdUpdatePrice 
      Caption         =   " ÕœÌÀ «·«”⁄«—"
      Height          =   375
      Left            =   2070
      TabIndex        =   57
      Top             =   7350
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox Text4 
      Height          =   555
      Left            =   7200
      TabIndex        =   54
      Text            =   "Text4"
      Top             =   5730
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox txtCountSalesReturn 
      Height          =   375
      Left            =   5220
      TabIndex        =   51
      Top             =   4020
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtCountSales 
      Height          =   375
      Left            =   5640
      TabIndex        =   50
      Top             =   3480
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtEndTime 
      Height          =   375
      Left            =   3390
      TabIndex        =   45
      Top             =   660
      Width           =   2205
   End
   Begin VB.TextBox txtStartTime 
      Height          =   375
      Left            =   3390
      TabIndex        =   44
      Top             =   210
      Width           =   2205
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   60
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   4170
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Height          =   675
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "FRMTRansferData3.frx":058A
      Top             =   3660
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«Œ »«— "
      Height          =   405
      Left            =   2100
      TabIndex        =   40
      Top             =   8880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   90
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   3390
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Frame Frame3 
      Caption         =   "«‰Ê«⁄ «·Õ—ﬂ« "
      Height          =   1035
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4590
      Width           =   4365
      Begin VB.CheckBox chkDefComItem 
         Caption         =   "”‰œ «· Ã„Ì⁄"
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   660
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkSalesOffers 
         Caption         =   "⁄—÷ «·”⁄—"
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox chkPay 
         Caption         =   "«·„œ›Ê⁄« "
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1410
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkRec 
         Caption         =   "«·„ﬁ»Ê÷« "
         Height          =   255
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   930
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "Õ—ﬂ«  «·Ê«—œ"
         Height          =   255
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1770
         Width           =   1635
      End
      Begin VB.CheckBox chkOut 
         Caption         =   "Õ—ﬂ«  «·’«œ—"
         Height          =   255
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   1305
      End
      Begin VB.CheckBox chkPurchaseReturn 
         Caption         =   "„— Ã⁄ «·„‘ —Ì« "
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkPurchase 
         Caption         =   "«·„‘ —Ì« "
         Height          =   255
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox chkSalesReturn 
         Caption         =   "„— Ã⁄ «·„»Ì⁄« "
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   150
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkSales 
         Caption         =   "«·„»Ì⁄« "
         Height          =   255
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   360
         Value           =   1  'Checked
         Width           =   1305
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ÕœÌÀ „·› «·«’‰«› Ê«·ÊÕœ« "
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   7410
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox DbName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   7530
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.ComboBox ServersName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   7440
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Data"
      Height          =   1455
      Left            =   0
      TabIndex        =   19
      Top             =   60
      Width           =   2145
      Begin VB.TextBox TxtServerDataBaseName 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Text            =   "byte"
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox DestinationServer 
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Text            =   "20.108.24.37,1433"
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server name"
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DBname"
         Height          =   375
         Left            =   -360
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame ServerData 
      Caption         =   "POS Data"
      Height          =   1935
      Left            =   0
      TabIndex        =   16
      Top             =   1470
      Width           =   2055
      Begin VB.ComboBox POSname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox TxtPOSDB 
         Height          =   375
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "Byte"
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox POSlServer 
         Height          =   375
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DBname"
         Height          =   375
         Left            =   -570
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server name"
         Height          =   375
         Left            =   -210
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "‰ﬁ· «·»Ì«‰« "
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5970
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -450
      Width           =   6135
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5280
         Top             =   3120
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "..."
         Height          =   375
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   5040
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TxtCHECKTIME 
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Text            =   "CHECKTIME Field Name"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox TxtUSERID 
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Text            =   "USERID Field Name"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox TxtTableName 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Text            =   "TableName"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDbPath 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "Database Path"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "‰ﬁ· «·»Ì«‰« "
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   3240
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DcTime 
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         Top             =   2520
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111017986
         CurrentDate     =   38784
      End
      Begin VB.Label LblInfo 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Update Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Date/Time Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Machhine Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Table Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "DB Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtLicense 
      Alignment       =   1  'Right Justify
      Height          =   1095
      Left            =   -840
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   7935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10380
      Top             =   6180
      Visible         =   0   'False
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
   Begin MSComCtl2.DTPicker dbRecordDate 
      Height          =   285
      Left            =   4860
      TabIndex        =   41
      Top             =   4530
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   41640
   End
   Begin VSFlex8UCtl.VSFlexGrid grd 
      Height          =   5130
      Left            =   6810
      TabIndex        =   55
      Top             =   540
      Width           =   6525
      _cx             =   11509
      _cy             =   9049
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FRMTRansferData3.frx":0590
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   0
      AccessibleName  =   "ReCostDet"
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.DTPicker txtToDate 
      Height          =   285
      Left            =   4860
      TabIndex        =   74
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker txtFromDate 
      Height          =   285
      Left            =   4860
      TabIndex        =   75
      Top             =   4860
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   41640
   End
   Begin VSFlex8UCtl.VSFlexGrid grdInfo 
      Height          =   2070
      Left            =   2220
      TabIndex        =   94
      Top             =   1080
      Width           =   4515
      _cx             =   7964
      _cy             =   3651
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FRMTRansferData3.frx":066D
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   0
      AccessibleName  =   "ReCostDet"
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1020
      TabIndex        =   91
      Top             =   6030
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ ⁄—Ê÷ «·”⁄— «·„‰ﬁÊ·…"
      Height          =   255
      Index           =   4
      Left            =   5310
      TabIndex        =   80
      Top             =   4470
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "√Ì«„ ·„ Ì „ ‰ﬁ·Â«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Index           =   3
      Left            =   7950
      TabIndex        =   56
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label lblWait 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì—ÃÏ «·«‰ Ÿ«— Ã«—Ì ‰ﬁ· «·»Ì«‰« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1245
      Left            =   210
      TabIndex        =   53
      Top             =   5670
      Visible         =   0   'False
      Width           =   5265
   End
   Begin VB.Label lblCount 
      Height          =   315
      Left            =   1410
      TabIndex        =   52
      Top             =   6090
      Width           =   1425
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·„—œÊœ«  «·„‰ﬁÊ·…"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   49
      Top             =   3720
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·›Ê« Ì— «·„‰ﬁÊ·…"
      Height          =   255
      Index           =   1
      Left            =   5370
      TabIndex        =   48
      Top             =   3240
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Êﬁ  «·‰Â«Ì…"
      Height          =   255
      Left            =   5760
      TabIndex        =   47
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Êﬁ  «·»œ«Ì…"
      Height          =   255
      Index           =   0
      Left            =   5730
      TabIndex        =   46
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "FRMTRansferData3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Dim CountItems As Long, CountSales As Long, CountSalesReturn As Long, CountPurchase As Long, CountPurchaseReturn As Long, mCounCountRec, CountSalesOfeers As Long
 Dim cProgress As ClsProgress
 Dim BolFrmLoaded As Boolean
Dim s As String
 Dim mNoteType As Integer
    Dim mSanadNo As Integer
Dim mUserId As Long
Dim mTimeStart As Date, mEndTime As Date, ActualDeliveryDate As Date, LatestDeliveryDate As Date
Dim mDBPOSName As String

Private mLastSQL As String
Private mLastStep As String

Private Sub CmdOpen_Click()
cd1.ShowOpen
 
txtDbPath.Text = cd1.FileName


End Sub
Function CopyIssueTtransaction(invoiceTransaction_ID As Double, invoiceNoteserial1 As String, Transaction_ID As Double, Transaction_Type As Double, issuenoteserial As String, issuenoteserial1 As String, SessionCode As String)
'////////////////////////////////////////copy   Transactions



  Dim Rs3 As ADODB.Recordset
  Dim rsDouble_Entry As ADODB.Recordset
  
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim mytext As String
    
 sql = " select * from Transactions    WHERE Transaction_ID =" & Transaction_ID
 
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
 Dim FromTransaction_ID As Double
 Dim FromBranchID As Integer
 Dim FromTransaction_Date As Date
 Dim FromNots As String
  Dim FromNots2 As String
   Dim fromTransaction_Serial As String
 Dim FromNoteseial1 As String
 Dim FromTransaction_Type As Integer
 
Dim BranchID As Integer
 ' Dim Transaction_ID As Double
  Dim Transaction_Date As Date
   Dim Transaction_Serial  As String
 Dim Nots As String
Dim Nots2 As String
'Dim Transaction_Type As Integer
Dim FromNoteId As Double
 'sales
    If Rs3.RecordCount > 0 Then
      
            For i = 1 To Rs3.RecordCount
             FromTransaction_Type = IIf(IsNull(Rs3("Transaction_Type").Value), 0, Rs3("Transaction_Type").Value)
               FromTransaction_ID = IIf(IsNull(Rs3("Transaction_ID").Value), 0, Rs3("Transaction_ID").Value)
                
               mUserId = Val(Rs3!userID & "")
               
               FromBranchID = IIf(IsNull(Rs3("BranchID").Value), 0, Rs3("BranchID").Value)
               fromTransaction_Serial = IIf(IsNull(Rs3("Transaction_Serial").Value), 0, Rs3("Transaction_Serial").Value)
        
              
               FromNoteSerial1 = IIf(IsNull(Rs3("Noteserial1").Value), 0, Rs3("Noteserial1").Value)
                FromNoteSerial = IIf(IsNull(Rs3("Noteserial").Value), 0, Rs3("Noteserial").Value)
                FromNoteId = IIf(IsNull(Rs3("NoteId").Value), 0, Rs3("NoteId").Value) ' —ﬁ„ ﬁÌœ «·”‰œ
                
               FromNots2 = IIf(IsNull(Rs3("Nots2").Value), 0, Rs3("Nots2").Value) '—ﬁ„ «·›« Ê—… «·«’·ÌÏ…
               FromTransaction_Date = IIf(IsNull(Rs3("Transaction_Date").Value), 0, Rs3("Transaction_Date").Value)
              
                      Dim FromEmp_ID As Double

 Dim FromStoreID As Double
Dim FromCusID As Double
               Dim FromBoxid As Double
            Dim PayMentType As Integer
               Dim BillBasedOn
           'Dim BillBasedOn As Integer
              Dim VATYou As Double
               Dim VAT As Double
               Dim FromUserID As Double
               Dim POSBillType As Double
               FromUserID = IIf(IsNull(Rs3("UserID").Value), 0, Rs3("UserID").Value)
               FromEmp_ID = IIf(IsNull(Rs3("Emp_ID").Value), 0, Rs3("Emp_ID").Value)
               FromStoreID = IIf(IsNull(Rs3("storeID").Value), 0, Rs3("storeID").Value)
               FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
               
               FromBoxid = IIf(IsNull(Rs3("Boxid").Value), 0, Rs3("Boxid").Value)
               POSBillType = IIf(IsNull(Rs3("POSBillType").Value), 0, Rs3("POSBillType").Value)
               
               
                FromPaymentType = IIf(IsNull(Rs3("PaymentType").Value), 0, Rs3("PaymentType").Value)
                FromBillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                FromVATYou = IIf(IsNull(Rs3("VATYou").Value), 0, Rs3("VATYou").Value)
                FromVAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
               '


              Transaction_Date = FromTransaction_Date
              Transaction_Type3 = FromTransaction_Type
              BranchID = FromBranchID
             Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
             Transaction_Serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=" & Transaction_Type & ""))
             If Transaction_Type = 19 Then
 
              NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, 10, 180, , 19, , , , , , FromUserID)
             ElseIf Transaction_Type = 20 Then
 
              NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, 9, 160, , 20, , , , , , FromUserID)
              
             End If
             
             
             NoteSerial = Notes_coding(FromBranchID, FromTransaction_Date)
             NoteId = CStr(new_id("Notes", "NoteID", "", True))
            TransactionComment = " ”‰œ „‰ﬁÊ· „‰ ﬁ«⁄œ…  " & POSname.Text & "   "
            TransactionComment = TransactionComment & "  —ﬁ„ «·”‰œ «·«’·Ì   " & FromNoteSerial1
             '" & ServerDb & "
             
 
              
'ÂÌœ— «·”‰œ
'*****************************************************************************************

'*****************************************************************************************
 sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transactions]  (    "
sql = sql & "  Transaction_ID,Transaction_Date, Transaction_Serial , Transaction_Type, PaymentType, CusID, StoreID, UserID, Emp_ID, BranchId, BoxID  "
sql = sql & " , BillBasedOn, VAT, VATYou, NoteSerial,NoteSerial1,NoteId,Copied,TransactionComment,closed,SessionCode,OldNoteSerial1,OldNoteSerial,OldNoteId,OldTransaction_ID)"
 
sql = sql & "   values (" & Transaction_ID & "," & SQLDate(Transaction_Date, True) & ", " & Transaction_Serial & "," & Transaction_Type & "," & FromPaymentType & "," & FromCusID & "," & FromStoreID & ",1," & FromEmp_ID & "," & FromBranchID & "," & FromBoxid
sql = sql & "," & FromBillBasedOn & "," & FromVAT & "," & FromVATYou & ",'" & NoteSerial & "','" & NoteSerial1 & "'," & NoteId & ",1,'" & TransactionComment & "',1,'" & SessionCode & "',"

sql = sql & "'" & FromNoteSerial1 & "' , "
sql = sql & "'" & FromNoteSerial & "' , " & FromNoteId & " , " & FromTransaction_ID & " )"

            '   fromTransaction_Serial
 
 
 mLastStep = "Insert missing items into remote server"
mLastSQL = sql

 Cn.Execute sql

Text2.Text = sql
      
      
     ' ›«’Ì· «·”‰œ
  
 
 sql = " select * from Transaction_Details   where  Transaction_ID=" & FromTransaction_ID
    Set rsDouble_Entry = New ADODB.Recordset
    '
   rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    Dim j As Double
    For j = 1 To rsDouble_Entry.RecordCount
    Item_ID = IIf(IsNull(rsDouble_Entry("Item_ID").Value), 0, rsDouble_Entry("Item_ID").Value)
     ItemCase = IIf(IsNull(rsDouble_Entry("ItemCase").Value), 0, rsDouble_Entry("ItemCase").Value)
      Quantity = IIf(IsNull(rsDouble_Entry("Quantity").Value), 0, rsDouble_Entry("Quantity").Value)
       Price = IIf(IsNull(rsDouble_Entry("Price").Value), 0, rsDouble_Entry("Price").Value)
        ItemDiscountType = IIf(IsNull(rsDouble_Entry("ItemDiscountType").Value), 0, rsDouble_Entry("ItemDiscountType").Value)
         ItemDiscount = IIf(IsNull(rsDouble_Entry("ItemDiscount").Value), 0, rsDouble_Entry("ItemDiscount").Value)
         ShowQty = IIf(IsNull(rsDouble_Entry("ShowQty").Value), 0, rsDouble_Entry("ShowQty").Value)
         showPrice = IIf(IsNull(rsDouble_Entry("showPrice").Value), 0, rsDouble_Entry("showPrice").Value)
         UnitID = IIf(IsNull(rsDouble_Entry("UnitId").Value), 0, rsDouble_Entry("UnitId").Value)
         ColorID = IIf(IsNull(rsDouble_Entry("ColorID").Value), 0, rsDouble_Entry("ColorID").Value)
         ItemSize = IIf(IsNull(rsDouble_Entry("ItemSize").Value), 0, rsDouble_Entry("ItemSize").Value)
         ClassId = IIf(IsNull(rsDouble_Entry("ClassId").Value), 0, rsDouble_Entry("ClassId").Value)
         
         
 
    sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transaction_Details]  (    "
sql = sql & "  Transaction_ID,  Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice,UnitId , ColorID, ItemSize, ClassId,SessionCode)"
 sql = sql & "   values (" & Transaction_ID & "," & Item_ID & ", " & ItemCase & "," & Quantity & "," & Price & "," & ItemDiscountType & "," & ItemDiscount & "," & ShowQty & "," & showPrice
 sql = sql & "," & UnitID & "," & ColorID & "," & ItemSize & "," & ClassId & "" & ",'" & SessionCode & "')"
 
 
  mLastStep = "Insert missing items into remote server"
mLastSQL = sql
           Cn.Execute sql
           rsDouble_Entry.MoveNext
    Next j
    
 
 
         
         
'ﬁÌœ «·”‰œ
  

sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,Transaction_ID,SessionCode)"
 sql = sql & " values( " & NoteId & ", " & SQLDate(Transaction_Date, True) & " , " & mNoteType & ", " & NoteSerial & ", " & NoteSerial1 & "," & BranchID & "," & Transaction_ID & ",'" & SessionCode & "')"
 
  mLastStep = "Insert missing items into remote server"
mLastSQL = sql

 Cn.Execute sql
' MsgBox "ﬁÌœ «·”‰œ"
 DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
 
 
 'Dim rsDouble_Entry As ADODB.Recordset
  Set rsDouble_Entry = New ADODB.Recordset
     sql = " select * from DOUBLE_ENTREY_VOUCHERS   where   Notes_ID=" & FromNoteId
   rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    Dim w As Double
    For w = 1 To rsDouble_Entry.RecordCount
    Account_Code = IIf(IsNull(rsDouble_Entry("Account_Code").Value), 0, rsDouble_Entry("Account_Code").Value)
    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
    Credit_Or_Debit = IIf(IsNull(rsDouble_Entry("Credit_Or_Debit").Value), 0, rsDouble_Entry("Credit_Or_Debit").Value)
    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
    Double_Entry_Vouchers_Description = IIf(IsNull(rsDouble_Entry("Double_Entry_Vouchers_Description").Value), 0, rsDouble_Entry("Double_Entry_Vouchers_Description").Value) & Chr(13) & "  ”‰œ ’—› " & TransactionComment
    'RecordDate = IIf(IsNull(rsDouble_Entry("RecordDate").Value), 0, rsDouble_Entry("RecordDate").Value)
    DEV_ID_Line_No = IIf(IsNull(rsDouble_Entry("DEV_ID_Line_No").Value), 0, rsDouble_Entry("DEV_ID_Line_No").Value)
    branch_id = IIf(IsNull(rsDouble_Entry("branch_id").Value), 0, rsDouble_Entry("branch_id").Value)
    sql = "  INSERT INTO [" & ServerDb & "].[dbo].[DOUBLE_ENTREY_VOUCHERS]([Double_Entry_Vouchers_ID], [DEV_ID_Line_No], [Account_Code], [Value], [Credit_Or_Debit], [Double_Entry_Vouchers_Description], [RecordDate], [Notes_ID] ,branch_id,UserID,Transaction_ID,SessionCode) "
    sql = sql & " values (  " & DEVID & ", " & DEV_ID_Line_No & ", '" & Account_Code & "', " & Value & ", " & Credit_Or_Debit & ", '" & Double_Entry_Vouchers_Description & "',  " & SQLDate(Transaction_Date, True) & ", " & NoteId & " ," & branch_id & ",1 ," & Transaction_ID & ",'" & SessionCode & "')"
  
  mLastStep = "Insert missing items into remote server"
mLastSQL = sql
  Cn.Execute sql


    rsDouble_Entry.MoveNext
    Next w
   
  
  
 
'*****************************************************************
'**********************************************************
 


  
         
         
        
     Next i
     
     
      End If
 
    Rs3.Close
  'Sql = Sql & "[" & POSDb & "].dbo.Transactions"
  '„‰⁄ «·‰ﬁ· „—… «Œ—Ì
  sql = "update   [" & POSDb & "].dbo.Transactions" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE   Transaction_ID =" & FromTransaction_ID
 
 mLastStep = "Insert missing items into remote server"
mLastSQL = sql
 POSConnection.Execute sql
 

  sql = "update   [" & POSDb & "].dbo.Transaction_Details" & "  set  Copied =1,SessionCode = '" & SessionCode & "' where Transaction_ID =" & FromTransaction_ID
   mLastStep = "Insert missing items into remote server"
mLastSQL = sql
 POSConnection.Execute sql

 
     StrSQL = "UPDATE  [" & ServerDb & "].dbo. Transactions SET NOTS=" & invoiceTransaction_ID & ",NOTS2= '" & invoiceNoteserial1 & "' ,SessionCode = '" & SessionCode & "' WHERE Transaction_ID=" & Transaction_ID
        
        mLastStep = "Insert missing items into remote server"
mLastSQL = sql
        Cn.Execute StrSQL
             StrSQL = "UPDATE  [" & ServerDb & "].dbo. Transactions SET NOTS=" & Transaction_ID & ",NOTS2= '" & NoteSerial1 & "',SessionCode = '" & SessionCode & "' WHERE Transaction_ID=" & invoiceTransaction_ID
              mLastStep = "Insert missing items into remote server"
mLastSQL = sql
        Cn.Execute StrSQL
        
        
End Function
'Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean
'
'On Error GoTo ErrTrap
''«” ›”«—
''ServerDb = TxtServerDataBaseName.Text
''wael
''ServerDb = DestinationServer
'' POSDb = TxtServerDataBaseName.Text
'
'
'ServerDb = TxtServerDataBaseName.Text
'MsgBox ServerDb
'     Set Cn = New ADODB.Connection
'    With Cn
'        .CommandTimeout = 5000
'        .CursorLocation = adUseServer
'        .ConnectionTimeout = 5000
'       If SysSQLServerType = 1 Then
''        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
''        "Persist Security Info=False;Initial Catalog=" & ServerDb & _
''        ";Data Source=" & SysSQLServerName & ";Port=1433"
'
'
''        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
''        "Persist Security Info=False;Initial Catalog=" & ServerDb & _
''        ";Data Source=" & DestinationServer & ";Port=1433"
'
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & ServerDb & ";Data Source=" & DestinationServer 'SysSQLServerName
'        ElseIf SysSQLServerType = 2 Then
'
'
'                 If SysSQLServerTypeTechnical = "0" Then
'                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'                    "Persist Security Info=False;Initial Catalog=" & ServerDb & _
'                    ";Data Source=" & SysSQLServerName & ";Port=51433"
'                    '";Data Source=" & ServerDb & ";Port=1433"
'
'                  Else
'                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & ServerDb & ";Data Source=" & SysSQLServerName 'SysSQLServerName
'                End If
'          End If
'
'.Open
'End With
'ConnectionFirst = True
'
'
''ServerDb = TxtServerDataBaseName.Text
''wael
'
'If IsLoad Then Exit Function
'POSDb = TxtPOSDB.Text
'POSServer = POSlServer.Text
'
'
'     Set POSConnection = New ADODB.Connection
'    With POSConnection
'        .CommandTimeout = 5000
'        .CursorLocation = adUseClient
'        .ConnectionTimeout = 5000
'       If SysSQLServerType = 1 Then
''        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
''        "Persist Security Info=False;Initial Catalog=" & POSDb & _
''        ";Data Source=" & POSServer
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
'        ElseIf SysSQLServerType = 2 Then
'
'
'                 If SysSQLServerTypeTechnical = "0" Then
'                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'                    "Persist Security Info=False;Initial Catalog=" & POSDb & _
'                    ";Data Source=" & POSServer & ";Port=51433"
'                    '";Data Source=" & ServerDb & ";Port=1433"
'
'                  Else
'                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
'                End If
'          End If
'
'.Open
'
'End With
'ConnectionFirst = True
'
'
'
'Dim mPosD  As String
'Dim mServerD  As String
'mPosD = "[" & POSlServer & "]" & ".Master.dbo."
'mServerD = "[" & SysSQLServerName & "]" & ".Master.dbo."
'
'Dim s As String
'Dim ss As String
'
'    s = " USE MASTER " & vbNewLine
'    s = s & " DECLARE @sql NVARCHAR(4000) " & vbNewLine
'
'    s = s & " DECLARE db_cursor CURSOR FOR " & vbNewLine
'    s = s & "         select 'sp_dropserver ''' + [srvName] + '''' from sysservers " & vbNewLine
'
'    s = s & "     OPEN db_cursor " & vbNewLine
'    s = s & "     FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine
'
'    s = s & "     WHILE @@FETCH_STATUS = 0 " & vbNewLine
'    s = s & "     BEGIN " & vbNewLine
'
'    s = s & "            EXEC (@sql) " & vbNewLine
'
'    s = s & "            FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine
'    s = s & "     End " & vbNewLine
'
'    s = s & "     Close db_cursor " & vbNewLine
'    s = s & "     DEALLOCATE db_cursor " & vbNewLine
'
'    ss = "     USE " & ServerDb & vbNewLine
'
'    Cn.Execute s & ss
'    ss = "USE " & POSDb & vbNewLine
'    POSConnection.Execute s & ss
'   'POSConnection.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123"
'Dim rsDummy As New ADODB.Recordset
''s = "select * from " & mServerD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
''rsDummy.Open s, Cn, adOpenStatic
''If rsDummy.EOF Then
''    Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
''   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
''End If
''rsDummy.Close
'
''s = "select * from sys.servers Where name Like '" & SysSQLServerName & "'"
'
'
''s = "select * from sys.servers Where name Like '" & POSServer & "'"
's = "select * from sysservers Where srvName Like '" & POSServer & "'"
'rsDummy.Open s, Cn, adOpenStatic
'If rsDummy.EOF Then
'    Cn.Execute "EXEC sp_addlinkedserver [" & POSServer & "]"
'   ' Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If
'
'
'
''s = "select * from " & mServerD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
's = "select * from sysservers Where srvName Like '" & SysSQLServerName & "'"
'rsDummy.Close
'rsDummy.Open s, Cn, adOpenStatic
'If rsDummy.EOF Then
'
'    Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'   ' POSConnection.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If
'
'
''rsDummy.Close
's = " Use Master "
'POSConnection.Execute s
'
''s = "select * from " & mPosD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
's = "select * from sysservers Where  datasource Like '" & SysSQLServerName & "' and  srvName Like 'RemoteServer10'"
'rsDummy.Close
'rsDummy.Open s, POSConnection, adOpenStatic
'If rsDummy.EOF Then
'
'
'   s = "EXEC sp_addlinkedserver " & _
'      "@server = 'RemoteServer10', " & _
'      "@srvproduct = '', " & _
'      "@provider = 'SQLNCLI', " & _
'      "@datasrc = '172.187.248.126,51433';"
'
'POSConnection.Execute s
'
'' ≈⁄œ«œ  ”ÃÌ· «·œŒÊ· ··‹ Linked Server
's = "EXEC sp_addlinkedsrvlogin " & _
'      "@rmtsrvname = 'RemoteServer10', " & _
'      "@useself = 'false', " & _
'      "@rmtuser = 'sa', " & _
'      "@rmtpassword = 'Admin@123';"
'
'POSConnection.Execute s
'
'
'
''    POSConnection.Execute " EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If
'
'rsDummy.Close
'
's = "select * from sysservers Where srvName Like '" & POSServer & "'"
'
'rsDummy.Open s, POSConnection, adOpenStatic
'If rsDummy.EOF Then
'
'    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If
'rsDummy.Close
'
'
'
''s = "select * from " & mPosD & "sysservers Where srvName Like '" & POSServer & "'"
''rsDummy.Open s, POSConnection, adOpenStatic
''If rsDummy.EOF Then
''
''    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
''   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
''End If
''rsDummy.Close
'
'
'POSConnection.Execute "SET NOCOUNT ON;"
'POSConnection.Execute "SET LOCK_TIMEOUT 5000;"   ' 5 ????? ?????? ??? ???
'POSConnection.Execute "SET XACT_ABORT ON;"
'
'' ?? ????? ?? Cn ????:
'Cn.Execute "SET NOCOUNT ON;"
'Cn.Execute "SET LOCK_TIMEOUT 5000;"
'Cn.Execute "SET XACT_ABORT ON;"
'
'
'Set rsDummy = New ADODB.Recordset
's = "Select * from [" & SysSQLServerName & "]." & ServerDb & ".dbo.TblOptions "
'rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
'If Not rsDummy.EOF Then
'    NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
'    StoreDigit = Val(rsDummy!StoreDigit & "")
'    BranchDigit = Val(rsDummy!BranchDigit & "")
'    IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
'    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'    InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
'    ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
'    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'    NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
'    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'    IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
'    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'
'End If
'
'rsDummy.Close
''
''s = "select * from sys.servers Where name Like '" & POSServer & "'"
''rsDummy.Open s, POSConnection, adOpenStatic
''If rsDummy.EOF Then
''    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
''   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
''End If
'
'
'
''Do While Not rsDummy.EOF
''
''
''    rsDummy.MoveNext
''Loop
'
'
'
'Exit Function
'ErrTrap:
'Text1 = Cn.ConnectionString
'Text2 = POSConnection.ConnectionString
'frmPopup.ShowMessage "Õÿ√ ›Ì «·« ’«·"
' ConnectionFirst = False
'
'
'End Function
'


'Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean
'
'On Error GoTo ErrTrap
'
'Dim LastSQL As String
'Dim Extra As String
'
''========================
'' Main DB Connection
''========================
'ServerDb = TxtServerDataBaseName.Text
''MsgBox ServerDb
'
'Set Cn = New ADODB.Connection
'With Cn
'    .CommandTimeout = 5000
'    .CursorLocation = adUseServer
'    .ConnectionTimeout = 5000
'DestinationServer = SysSQLServerName
'    If SysSQLServerType = 1 Then
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                            ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                            ";Initial Catalog=" & ServerDb & _
'                            ";Data Source=" & DestinationServer
'    ElseIf SysSQLServerType = 2 Then
'        If SysSQLServerTypeTechnical = "0" Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'                                "Persist Security Info=False;Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & SysSQLServerName & ";Port=51433"
'        Else
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & SysSQLServerName
'        End If
'    End If
'
'    On Error Resume Next
'    .Errors.Clear
'    On Error GoTo ErrTrap
''MsgBox "Password " & SysSQLServerUserpassword & " Bb = " & ServerDb & " Data source " & SysSQLServerName & "SysSqlType = " & SysSQLServerType
'    .Open
'   ' MsgBox .ConnectionString
'
'   ' MsgBox "Status" & .State
'
'End With
'
'ConnectionFirst = True
'
'If IsLoad Then Exit Function
'
''========================
'' POS DB Connection
''========================
'POSDb = TxtPOSDB.Text
'POSServer = POSlServer.Text
'
'Set POSConnection = New ADODB.Connection
'With POSConnection
'    .CommandTimeout = 5000
'    .CursorLocation = adUseClient
'    .ConnectionTimeout = 5000
'
'    If SysSQLServerType = 1 Then
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                            ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                            ";Initial Catalog=" & POSDb & _
'                            ";Data Source=" & POSServer
'    ElseIf SysSQLServerType = 2 Then
'        If SysSQLServerTypeTechnical = "0" Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'                                "Persist Security Info=False;Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer & ";Port=51433"
'        Else
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer
'        End If
'    End If
'
'  '  On Error Resume Next
'    .Errors.Clear
'    On Error GoTo ErrTrap
''MsgBox "Befor Connection 2"
''MsgBox "Password " & SysSQLServerUserpassword & " Bb = " & POSDb & " Data source " & POSServer & "SysSqlType = " & SysSQLServerType
'    .Open
''        MsgBox .ConnectionString
'
''    MsgBox "Status" & .State
'End With
'
'ConnectionFirst = True
'
''========================
'' Your existing SQL work
''========================
'Dim mPosD As String
'Dim mServerD As String
'mPosD = "[" & POSlServer & "]" & ".Master.dbo."
'mServerD = "[" & SysSQLServerName & "]" & ".Master.dbo."
'
'Dim s As String
'Dim ss As String
'
's = " USE MASTER " & vbNewLine
's = s & " DECLARE @sql NVARCHAR(4000) " & vbNewLine
's = s & " DECLARE db_cursor CURSOR FOR " & vbNewLine
's = s & "         select 'sp_dropserver ''' + [srvName] + '''' from sysservers " & vbNewLine
's = s & "     OPEN db_cursor " & vbNewLine
's = s & "     FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine
's = s & "     WHILE @@FETCH_STATUS = 0 " & vbNewLine
's = s & "     BEGIN " & vbNewLine
's = s & "            EXEC (@sql) " & vbNewLine
's = s & "            FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine
's = s & "     End " & vbNewLine
's = s & "     Close db_cursor " & vbNewLine
's = s & "     DEALLOCATE db_cursor " & vbNewLine
'
'ss = "     USE " & ServerDb & vbNewLine
'
'LastSQL = s & ss
'
' mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'
'Cn.Execute LastSQL
'
'ss = "USE " & POSDb & vbNewLine
'LastSQL = s & ss
'
' mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'
'POSConnection.Execute LastSQL
'
'Dim rsDummy As New ADODB.Recordset
'
's = "select * from sysservers Where srvName Like '" & POSServer & "'"
'LastSQL = s
'rsDummy.Open s, Cn, adOpenStatic
'If rsDummy.EOF Then
'    LastSQL = "EXEC sp_addlinkedserver [" & POSServer & "]"
'     mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    Cn.Execute LastSQL
'End If
'
's = "select * from sysservers Where srvName Like '" & SysSQLServerName & "'"
'LastSQL = s
'rsDummy.Close
'rsDummy.Open s, Cn, adOpenStatic
'If rsDummy.EOF Then
'    LastSQL = "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'     mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    Cn.Execute LastSQL
'End If
'
'LastSQL = "Use Master"
'
' mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'POSConnection.Execute LastSQL
'
''s = "select * from sysservers Where  datasource Like '" & SysSQLServerName & "' and  srvName Like 'RemoteServer10'"
''LastSQL = s
''rsDummy.Close
''rsDummy.Open s, POSConnection, adOpenStatic
''If rsDummy.EOF Then
''
'''    LastSQL = "EXEC sp_addlinkedserver " & _
'''              "@server = 'RemoteServer10', " & _
'''              "@srvproduct = '', " & _
'''              "@provider = 'SQLNCLI', " & _
'''              "@datasrc = '172.187.248.126,51433';"
'''    POSConnection.Execute LastSQL
'''
'''    LastSQL = "EXEC sp_addlinkedsrvlogin " & _
'''              "@rmtsrvname = 'RemoteServer10', " & _
'''              "@useself = 'false', " & _
'''              "@rmtuser = 'sa', " & _
'''              "@rmtpassword = 'Admin@123';"
'''    POSConnection.Execute LastSQL
''
''LastSQL = "EXEC master.dbo.sp_addlinkedserver " & _
''          "@server = N'RemoteServer10', " & _
''          "@srvproduct = N'', " & _
''          "@provider = N'MSOLEDBSQL', " & _
''          "@datasrc = N'172.187.248.126,51433';"
''POSConnection.Execute LastSQL
''
''LastSQL = "EXEC master.dbo.sp_addlinkedsrvlogin " & _
''          "@rmtsrvname = N'RemoteServer10', " & _
''          "@useself = N'False', " & _
''          "@locallogin = NULL, " & _
''          "@rmtuser = N'sa', " & _
''          "@rmtpassword = N'Admin@123';"
''POSConnection.Execute LastSQL
''End If
''LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'data access', @optvalue=N'true';"
''POSConnection.Execute LastSQL
''
''LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc', @optvalue=N'true';"
''POSConnection.Execute LastSQL
''
''LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc out', @optvalue=N'true';"
''POSConnection.Execute LastSQL
'Dim rsChk As New ADODB.Recordset
'
'LastSQL = "SELECT 1 FROM master.dbo.sysservers WHERE srvname = 'RemoteServer10'"
'rsChk.Open LastSQL, POSConnection, adOpenStatic, adLockReadOnly
'
'If rsChk.EOF Then
'    rsChk.Close
'
'    LastSQL = "EXEC master.dbo.sp_addlinkedserver " & _
'              "@server = N'RemoteServer10', " & _
'              "@srvproduct = N'', " & _
'              "@provider = N'MSOLEDBSQL', " & _
'              "@datasrc = N'172.187.248.126,51433';"
'               mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_addlinkedsrvlogin " & _
'              "@rmtsrvname = N'RemoteServer10', " & _
'              "@useself = N'False', " & _
'              "@locallogin = NULL, " & _
'              "@rmtuser = N'sa', " & _
'              "@rmtpassword = N'Admin@123';"
'               mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'data access', @optvalue=N'true';"
'     mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc', @optvalue=N'true';"
'     mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc out', @optvalue=N'true';"
'     mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'Else
'    rsChk.Close
'End If
'
'Set rsChk = Nothing
'rsDummy.Close
'
's = "select * from sysservers Where srvName Like '" & POSServer & "'"
'LastSQL = s
' mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'rsDummy.Open s, POSConnection, adOpenStatic
'If rsDummy.EOF Then
'    LastSQL = "EXEC sp_addlinkedserver [" & POSServer & "]"
'     mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'End If
'rsDummy.Close
'
'LastSQL = "SET NOCOUNT ON;"
'POSConnection.Execute LastSQL
'LastSQL = "SET LOCK_TIMEOUT 5000;"
'POSConnection.Execute LastSQL
'LastSQL = "SET XACT_ABORT ON;"
'POSConnection.Execute LastSQL
'
'LastSQL = "SET NOCOUNT ON;"
'Cn.Execute LastSQL
'LastSQL = "SET LOCK_TIMEOUT 5000;"
'Cn.Execute LastSQL
'LastSQL = "SET XACT_ABORT ON;"
' mLastStep = "Insert missing items into remote server"
'mLastSQL = LastSQL
'Cn.Execute LastSQL
'
'Set rsDummy = New ADODB.Recordset
's = "Select * from [" & SysSQLServerName & "]." & ServerDb & ".dbo.TblOptions "
'LastSQL = s
'rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
'If Not rsDummy.EOF Then
'    NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
'    StoreDigit = Val(rsDummy!StoreDigit & "")
'    BranchDigit = Val(rsDummy!BranchDigit & "")
'    IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
'    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'    InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
'    ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
'    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'    NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
'    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'    IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
'    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'End If
'rsDummy.Close
'
'Exit Function
'
'ErrTrap:
'    ' ⁄‘«‰ „«ÌÕ’·‘ Error ÃÊÂ «· ErrorHandler ·Ê Cn/POSConnection „‘ „ ⁄„·Ì‰
'    On Error Resume Next
'
'    Extra = "ServerDb=" & ServerDb & vbCrLf & _
'            "POSDb=" & POSDb & vbCrLf & _
'            "DestinationServer=" & DestinationServer & vbCrLf & _
'            "SysSQLServerName=" & SysSQLServerName & vbCrLf & _
'            "POSServer=" & POSServer & vbCrLf & _
'            "SysSQLServerType=" & SysSQLServerType & vbCrLf & _
'            "SysSQLServerTypeTechnical=" & SysSQLServerTypeTechnical & vbCrLf & _
'            "Cn.ConnectionString=" & IIf(Cn Is Nothing, "", Cn.ConnectionString) & vbCrLf & _
'            "POS.ConnectionString=" & IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    LogErrDetailed "ConnectionFirst", Err, Erl, Extra, Cn, POSConnection, LastSQL
'
'    Text1 = IIf(Cn Is Nothing, "", Cn.ConnectionString)
'    Text2 = IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    frmPopup.ShowMessage "Õÿ√ ›Ì «·« ’«·"
'
'    ConnectionFirst = False
'End Function

'Public Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean
'
'    On Error GoTo ErrTrap
'
'    Dim LastSQL As String
'    Dim Extra As String
'
'    Dim rsDummy As ADODB.Recordset
'    Dim rsChk As ADODB.Recordset
'
'    Dim mRemoteDataSrc As String
'    Dim mRemoteProvider As String
'
'    ConnectionFirst = False
'    LastSQL = ""
'    Extra = ""
'
'    ServerDb = Trim$(TxtServerDataBaseName.Text & "")
'    POSDb = Trim$(TxtPOSDB.Text & "")
'    POSServer = Trim$(POSlServer.Text & "")
'    DestinationServer = Trim$(SysSQLServerName & "")
'
'    If Len(ServerDb) = 0 Then
'        frmPopup.ShowMessage "«”„ ﬁ«⁄œ… «·»Ì«‰«  «·—∆Ì”Ì… ›«—€"
'        Exit Function
'    End If
'
'    If Not IsLoad Then
'        If Len(POSDb) = 0 Then
'            frmPopup.ShowMessage "«”„ ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ… ›«—€"
'            Exit Function
'        End If
'
'        If Len(POSServer) = 0 Then
'            frmPopup.ShowMessage "«”„ ”Ì—›— «·‰ﬁÿ… ›«—€"
'            Exit Function
'        End If
'    End If
'
'    '========================================================
'    '  ‰ŸÌ› « ’«·«  ”«»ﬁ… ≈‰ ÊÃœ 
'    '========================================================
'    On Error Resume Next
'
'    If Not Cn Is Nothing Then
'        If Cn.State <> adStateClosed Then Cn.Close
'    End If
'    Set Cn = Nothing
'
'    If Not POSConnection Is Nothing Then
'        If POSConnection.State <> adStateClosed Then POSConnection.Close
'    End If
'    Set POSConnection = Nothing
'
'    On Error GoTo ErrTrap
'
'    '========================================================
'    ' Main DB Connection
'    '========================================================
'    Set Cn = New ADODB.Connection
'
'    With Cn
'        .CommandTimeout = 300
'        .CursorLocation = adUseServer
'        .ConnectionTimeout = 30
'
'        If SysSQLServerType = 1 Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & DestinationServer
'
'        ElseIf SysSQLServerType = 2 Then
'
'            If SysSQLServerTypeTechnical = "0" Then
'                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
'                                    "Persist Security Info=False;" & _
'                                    "Initial Catalog=" & ServerDb & _
'                                    ";Data Source=" & SysSQLServerName & ",51433"
'            Else
'                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                    ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                    ";Initial Catalog=" & ServerDb & _
'                                    ";Data Source=" & SysSQLServerName
'            End If
'        Else
'            Err.Raise vbObjectError + 9101, , "ﬁÌ„… SysSQLServerType €Ì— ’ÕÌÕ…"
'        End If
'
'        .Errors.Clear
'        LastSQL = ".Open Main Connection"
'        mLastStep = "Open Main Connection"
'        mLastSQL = .ConnectionString
'        .Open
'    End With
'
'    If IsLoad Then
'        ConnectionFirst = True
'        Exit Function
'    End If
'
'    '========================================================
'    ' POS DB Connection
'    '========================================================
'    Set POSConnection = New ADODB.Connection
'
'    With POSConnection
'        .CommandTimeout = 300
'        .CursorLocation = adUseClient
'        .ConnectionTimeout = 30
'
'        If SysSQLServerType = 1 Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer
'
'        ElseIf SysSQLServerType = 2 Then
'
'            If SysSQLServerTypeTechnical = "0" Then
'                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
'                                    "Persist Security Info=False;" & _
'                                    "Initial Catalog=" & POSDb & _
'                                    ";Data Source=" & POSServer & ",51433"
'            Else
'                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                    ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                    ";Initial Catalog=" & POSDb & _
'                                    ";Data Source=" & POSServer
'            End If
'        Else
'            Err.Raise vbObjectError + 9102, , "ﬁÌ„… SysSQLServerType €Ì— ’ÕÌÕ…"
'        End If
'
'        .Errors.Clear
'        LastSQL = ".Open POS Connection"
'        mLastStep = "Open POS Connection"
'        mLastSQL = .ConnectionString
'        .Open
'    End With
'
'    '========================================================
'    ' Session settings
'    '========================================================
'    LastSQL = "SET NOCOUNT ON;"
'    mLastStep = "Session Settings Main"
'    mLastSQL = LastSQL
'    Cn.Execute LastSQL, , adExecuteNoRecords
'
'    LastSQL = "SET LOCK_TIMEOUT 5000;"
'    mLastSQL = LastSQL
'    Cn.Execute LastSQL, , adExecuteNoRecords
'
'    LastSQL = "SET XACT_ABORT ON;"
'    mLastSQL = LastSQL
'    Cn.Execute LastSQL, , adExecuteNoRecords
'
'    LastSQL = "SET NOCOUNT ON;"
'    mLastStep = "Session Settings POS"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'    LastSQL = "SET LOCK_TIMEOUT 5000;"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'    LastSQL = "SET XACT_ABORT ON;"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'    '========================================================
'    '  √ﬂœ „‰ ÊÃÊœ linked server »«”„ POSServer ⁄·Ï «·”Ì—›— «·—∆Ì”Ì
'    '========================================================
'    Set rsDummy = New ADODB.Recordset
'
'    LastSQL = "SELECT 1 FROM master.sys.servers WHERE name = N'" & Replace(POSServer, "'", "''") & "'"
'    mLastStep = "Check Linked Server on Main for POSServer"
'    mLastSQL = LastSQL
'    rsDummy.Open LastSQL, Cn, adOpenStatic, adLockReadOnly
'
'    If rsDummy.EOF Then
'        rsDummy.Close
'
'        LastSQL = "EXEC master.dbo.sp_addlinkedserver @server = N'" & Replace(POSServer, "'", "''") & "'"
'        mLastStep = "Create Linked Server on Main for POSServer"
'        mLastSQL = LastSQL
'        Cn.Execute LastSQL, , adExecuteNoRecords
'    Else
'        rsDummy.Close
'    End If
'
'    '========================================================
'    '  √ﬂœ „‰ ÊÃÊœ linked server »«”„ SysSQLServerName ⁄·Ï «·”Ì—›— «·—∆Ì”Ì
'    '========================================================
'    LastSQL = "SELECT 1 FROM master.sys.servers WHERE name = N'" & Replace(SysSQLServerName, "'", "''") & "'"
'    mLastStep = "Check Linked Server on Main for SysSQLServerName"
'    mLastSQL = LastSQL
'    rsDummy.Open LastSQL, Cn, adOpenStatic, adLockReadOnly
'
'    If rsDummy.EOF Then
'        rsDummy.Close
'
'        LastSQL = "EXEC master.dbo.sp_addlinkedserver @server = N'" & Replace(SysSQLServerName, "'", "''") & "'"
'        mLastStep = "Create Linked Server on Main for SysSQLServerName"
'        mLastSQL = LastSQL
'        Cn.Execute LastSQL, , adExecuteNoRecords
'    Else
'        rsDummy.Close
'    End If
'
'    '========================================================
'    '  √ﬂœ „‰ ÊÃÊœ RemoteServer10 ⁄·Ï «·‰ﬁÿ…
'    '========================================================
'    Set rsChk = New ADODB.Recordset
'
'    LastSQL = "SELECT 1 FROM master.sys.servers WHERE name = N'RemoteServer10'"
'    mLastStep = "Check RemoteServer10 on POS"
'    mLastSQL = LastSQL
'    rsChk.Open LastSQL, POSConnection, adOpenStatic, adLockReadOnly
'
'    If rsChk.EOF Then
'        rsChk.Close
'
'        If SysSQLServerTypeTechnical = "0" Then
'            mRemoteDataSrc = SysSQLServerName & ",51433"
'        Else
'            mRemoteDataSrc = SysSQLServerName
'        End If
'
'        ' „Â„:
'        ' €Ì¯— «·‹ provider Â‰« ≈–« «·ÃÂ«“ ⁄‰œﬂ ·« ÌÕ ÊÌ MSOLEDBSQL
'        ' «·»œÌ· «·√ﬁ—» €«·»«: SQLNCLI11
'        mRemoteProvider = "SQLOLEDB"
'
'        LastSQL = "EXEC master.dbo.sp_addlinkedserver " & _
'                  "@server = N'RemoteServer10', " & _
'                  "@srvproduct = N'', " & _
'                  "@provider = N'" & mRemoteProvider & "', " & _
'                  "@datasrc = N'" & Replace(mRemoteDataSrc, "'", "''") & "';"
'        mLastStep = "Create RemoteServer10 on POS"
'        mLastSQL = LastSQL
'        POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'        LastSQL = "EXEC master.dbo.sp_addlinkedsrvlogin " & _
'                  "@rmtsrvname = N'RemoteServer10', " & _
'                  "@useself = N'False', " & _
'                  "@locallogin = NULL, " & _
'                  "@rmtuser = N'" & Replace(SysSQLServerUserId, "'", "''") & "', " & _
'                  "@rmtpassword = N'" & Replace(SysSQLServerUserpassword, "'", "''") & "';"
'        mLastStep = "Create RemoteServer10 Login on POS"
'        mLastSQL = LastSQL
'        POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'        LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'data access', @optvalue=N'true';"
'        mLastStep = "Enable data access on RemoteServer10"
'        mLastSQL = LastSQL
'        POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'        LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc', @optvalue=N'true';"
'        mLastStep = "Enable rpc on RemoteServer10"
'        mLastSQL = LastSQL
'        POSConnection.Execute LastSQL, , adExecuteNoRecords
'
'        LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc out', @optvalue=N'true';"
'        mLastStep = "Enable rpc out on RemoteServer10"
'        mLastSQL = LastSQL
'        POSConnection.Execute LastSQL, , adExecuteNoRecords
'    Else
'        rsChk.Close
'    End If
'
'    '========================================================
'    ' «Œ »«—Ì«: linked server »”Ìÿ »«”„ POSServer ⁄·Ï «·‰ﬁÿ… ·Ê „Õ «ÃÂ
'    '========================================================
'    LastSQL = "SELECT 1 FROM master.sys.servers WHERE name = N'" & Replace(POSServer, "'", "''") & "'"
'    mLastStep = "Check POSServer Linked Server on POS"
'    mLastSQL = LastSQL
'    rsDummy.Open LastSQL, POSConnection, adOpenStatic, adLockReadOnly
'
'    If rsDummy.EOF Then
'        rsDummy.Close
'
'        LastSQL = "EXEC master.dbo.sp_addlinkedserver @server = N'" & Replace(POSServer, "'", "''") & "'"
'        mLastStep = "Create POSServer Linked Server on POS"
'        mLastSQL = LastSQL
'        POSConnection.Execute LastSQL, , adExecuteNoRecords
'    Else
'        rsDummy.Close
'    End If
'
'    '========================================================
'    '  Õ„Ì· «·ŒÌ«—«  „‰ «·”Ì—›— «·—∆Ì”Ì
'    '========================================================
'    LastSQL = "SELECT * FROM dbo.TblOptions"
'    mLastStep = "Load TblOptions"
'    mLastSQL = LastSQL
'    rsDummy.Open LastSQL, Cn, adOpenKeyset, adLockReadOnly
'
'    If Not rsDummy.EOF Then
'        NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
'        StoreDigit = Val(rsDummy!StoreDigit & "")
'        BranchDigit = Val(rsDummy!BranchDigit & "")
'        IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
'        ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'        InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
'        ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
'        AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'        ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'        AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'        NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
'        JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'        IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
'        JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'    End If
'    rsDummy.Close
'
'    Set rsChk = Nothing
'    Set rsDummy = Nothing
'
'    ConnectionFirst = True
'    Exit Function
'
'ErrTrap:
'    On Error Resume Next
'
'    Extra = "ServerDb=" & ServerDb & vbCrLf & _
'            "POSDb=" & POSDb & vbCrLf & _
'            "DestinationServer=" & DestinationServer & vbCrLf & _
'            "SysSQLServerName=" & SysSQLServerName & vbCrLf & _
'            "POSServer=" & POSServer & vbCrLf & _
'            "SysSQLServerType=" & SysSQLServerType & vbCrLf & _
'            "SysSQLServerTypeTechnical=" & SysSQLServerTypeTechnical & vbCrLf & _
'            "Cn.ConnectionString=" & IIf(Cn Is Nothing, "", Cn.ConnectionString) & vbCrLf & _
'            "POS.ConnectionString=" & IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString) & vbCrLf & _
'            "LastStep=" & mLastStep
'
'    LogErrDetailed "ConnectionFirst", Err, Erl, Extra, Cn, POSConnection, LastSQL
'
'    Text1 = IIf(Cn Is Nothing, "", Cn.ConnectionString)
'    Text2 = IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    frmPopup.ShowMessage BuildErrMsg(Cn, "ConnectionFirst", LastSQL, "Œÿ√ ›Ì «·« ’«· √Ê ≈⁄œ«œ «·‹ Linked Server")
'
'    ConnectionFirst = False
'End Function

Private Function DeleteLinkedServer()
 

    
    
End Function

'Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean
'
'On Error GoTo ErrTrap
'
'Dim LastSQL As String
'Dim Extra As String
'Dim s As String
'
'Dim rsDummy As ADODB.Recordset
'Dim rsChk As ADODB.Recordset
'
'Dim providerName As String
'Dim dataSourceName As String
'Dim needRecreateRemote As Boolean
'Dim expectedDataSource As String
'
''========================
'' Main DB Connection
''========================
'ServerDb = Trim(TxtServerDataBaseName.Text)
'DestinationServer = Trim(SysSQLServerName)
'
'Set Cn = New ADODB.Connection
'With Cn
'    .CommandTimeout = 5000
'    .CursorLocation = adUseServer
'    .ConnectionTimeout = 5000
'
'    If SysSQLServerType = 1 Then
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                            ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                            ";Initial Catalog=" & ServerDb & _
'                            ";Data Source=" & DestinationServer
'    ElseIf SysSQLServerType = 2 Then
'        If SysSQLServerTypeTechnical = "0" Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
'                                "Persist Security Info=False;Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & DestinationServer & ";Port=51433"
'        Else
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & DestinationServer
'        End If
'    End If
'
'    .Errors.Clear
'    .Open
'End With
'
'ConnectionFirst = True
'If IsLoad Then Exit Function
'
''========================
'' POS DB Connection
''========================
'POSDb = Trim(TxtPOSDB.Text)
'POSServer = Trim(POSlServer.Text)
'
'Set POSConnection = New ADODB.Connection
'With POSConnection
'    .CommandTimeout = 5000
'    .CursorLocation = adUseClient
'    .ConnectionTimeout = 5000
'
'    If SysSQLServerType = 1 Then
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                            ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                            ";Initial Catalog=" & POSDb & _
'                            ";Data Source=" & POSServer
'    ElseIf SysSQLServerType = 2 Then
'        If SysSQLServerTypeTechnical = "0" Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
'                                "Persist Security Info=False;Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer & ";Port=51433"
'        Else
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer
'        End If
'    End If
'
'    .Errors.Clear
'    .Open
'End With
'
'ConnectionFirst = True
'
''=========================================================
'' ÷»ÿ RemoteServer10 ›ﬁÿ - »œÊ‰ Õ–› √Ì Linked Servers √Œ—Ï
'' Ê»œÊ‰ sp_testlinkedserver Õ Ï ·« Ì⁄ÿ· «·œ‰Ì«
''=========================================================
'expectedDataSource = Trim(SysSQLServerName) & ",51433"
'
'Set rsChk = New ADODB.Recordset
'providerName = ""
'dataSourceName = ""
'needRecreateRemote = False
'
'LastSQL = "SELECT TOP 1 provider, data_source FROM master.sys.servers WHERE name = 'RemoteServer10'"
'rsChk.Open LastSQL, POSConnection, adOpenStatic, adLockReadOnly
'
'If rsChk.EOF Then
'    needRecreateRemote = True
'Else
'    providerName = Trim(rsChk("provider").Value & "")
'    dataSourceName = Trim(rsChk("data_source").Value & "")
'
'    If UCase$(providerName) <> "SQLOLEDB" Then
'        needRecreateRemote = True
'    End If
'
'    If UCase$(Replace(dataSourceName, " ", "")) <> UCase$(Replace(expectedDataSource, " ", "")) Then
'        needRecreateRemote = True
'    End If
'End If
'
'If rsChk.State = adStateOpen Then rsChk.Close
'Set rsChk = Nothing
'
'If needRecreateRemote = True Then
'
'    On Error Resume Next
'    LastSQL = "EXEC master.dbo.sp_dropserver @server=N'RemoteServer10', @droplogins='droplogins';"
'    POSConnection.Execute LastSQL
'    On Error GoTo ErrTrap
'
'    LastSQL = "EXEC master.dbo.sp_addlinkedserver " & _
'              "@server = N'RemoteServer10', " & _
'              "@srvproduct = N'', " & _
'              "@provider = N'SQLOLEDB', " & _
'              "@datasrc = N'" & Replace(expectedDataSource, "'", "''") & "';"
'    mLastStep = "Create/Recreate RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_addlinkedsrvlogin " & _
'              "@rmtsrvname = N'RemoteServer10', " & _
'              "@useself = N'False', " & _
'              "@locallogin = NULL, " & _
'              "@rmtuser = N'" & Replace(SysSQLServerUserId, "'", "''") & "', " & _
'              "@rmtpassword = N'" & Replace(SysSQLServerUserpassword, "'", "''") & "';"
'    mLastStep = "Create login mapping for RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'data access', @optvalue=N'true';"
'    mLastStep = "Enable data access for RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc', @optvalue=N'true';"
'    mLastStep = "Enable rpc for RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'rpc out', @optvalue=N'true';"
'    mLastStep = "Enable rpc out for RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'connect timeout', @optvalue=N'20';"
'    mLastStep = "Set connect timeout for RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'
'    LastSQL = "EXEC master.dbo.sp_serveroption @server=N'RemoteServer10', @optname=N'query timeout', @optvalue=N'0';"
'    mLastStep = "Set query timeout for RemoteServer10"
'    mLastSQL = LastSQL
'    POSConnection.Execute LastSQL
'End If
'
''========================
'' Session options
''========================
'LastSQL = "SET NOCOUNT ON;"
'POSConnection.Execute LastSQL
'LastSQL = "SET LOCK_TIMEOUT 5000;"
'POSConnection.Execute LastSQL
'LastSQL = "SET XACT_ABORT ON;"
'POSConnection.Execute LastSQL
'
'LastSQL = "SET NOCOUNT ON;"
'Cn.Execute LastSQL
'LastSQL = "SET LOCK_TIMEOUT 5000;"
'Cn.Execute LastSQL
'LastSQL = "SET XACT_ABORT ON;"
'mLastStep = "Set XACT_ABORT ON for main connection"
'mLastSQL = LastSQL
'Cn.Execute LastSQL
'
''========================
'' Read options from central server
''========================
'Set rsDummy = New ADODB.Recordset
's = "SELECT * FROM TblOptions"
'LastSQL = s
'rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
'
'If Not rsDummy.EOF Then
'    NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
'    StoreDigit = Val(rsDummy!StoreDigit & "")
'    BranchDigit = Val(rsDummy!BranchDigit & "")
'    IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
'    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'    InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
'    ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
'    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'    NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
'    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'    IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
'End If
'
'rsDummy.Close
'Set rsDummy = Nothing
'
'Exit Function
'
'ErrTrap:
'    On Error Resume Next
'
'    Extra = "ServerDb=" & ServerDb & vbCrLf & _
'            "POSDb=" & POSDb & vbCrLf & _
'            "DestinationServer=" & DestinationServer & vbCrLf & _
'            "SysSQLServerName=" & SysSQLServerName & vbCrLf & _
'            "POSServer=" & POSServer & vbCrLf & _
'            "SysSQLServerType=" & SysSQLServerType & vbCrLf & _
'            "SysSQLServerTypeTechnical=" & SysSQLServerTypeTechnical & vbCrLf & _
'            "Cn.ConnectionString=" & IIf(Cn Is Nothing, "", Cn.ConnectionString) & vbCrLf & _
'            "POS.ConnectionString=" & IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    LogErrDetailed "ConnectionFirst", Err, Erl, Extra, Cn, POSConnection, LastSQL
'
'    Text1 = IIf(Cn Is Nothing, "", Cn.ConnectionString)
'    Text2 = IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    frmPopup.ShowMessage "ÕœÀ Œÿ√ ›Ì «·« ’«·"
'    ConnectionFirst = False
'End Function '

'Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean
'
'On Error GoTo ErrTrap
'
'Dim LastSQL As String
'Dim Extra As String
'Dim s As String
'Dim rsDummy As ADODB.Recordset
'
''========================
'' Main DB Connection
''========================
'ServerDb = Trim(TxtServerDataBaseName.Text)
'DestinationServer = Trim(SysSQLServerName)
'
'Set Cn = New ADODB.Connection
'With Cn
'    .CommandTimeout = 5000
'    .CursorLocation = adUseServer
'    .ConnectionTimeout = 5000
'
'    If SysSQLServerType = 1 Then
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                            ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                            ";Initial Catalog=" & ServerDb & _
'                            ";Data Source=" & DestinationServer
'    ElseIf SysSQLServerType = 2 Then
'        If SysSQLServerTypeTechnical = "0" Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
'                                "Persist Security Info=False;Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & DestinationServer & ";Port=51433"
'        Else
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & ServerDb & _
'                                ";Data Source=" & DestinationServer
'        End If
'    End If
'
'    .Errors.Clear
'    .Open
'End With
'
'ConnectionFirst = True
'
'If IsLoad Then Exit Function
'
''========================
'' POS DB Connection
''========================
'POSDb = Trim(TxtPOSDB.Text)
'POSServer = Trim(POSlServer.Text)
'
'Set POSConnection = New ADODB.Connection
'With POSConnection
'    .CommandTimeout = 5000
'    .CursorLocation = adUseClient
'    .ConnectionTimeout = 5000
'
'    If SysSQLServerType = 1 Then
'        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                            ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                            ";Initial Catalog=" & POSDb & _
'                            ";Data Source=" & POSServer
'    ElseIf SysSQLServerType = 2 Then
'        If SysSQLServerTypeTechnical = "0" Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
'                                "Persist Security Info=False;Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer & ";Port=51433"
'        Else
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & POSDb & _
'                                ";Data Source=" & POSServer
'        End If
'    End If
'
'    .Errors.Clear
'    .Open
'End With
'
'ConnectionFirst = True
'
''========================
'' Session options
''========================
'LastSQL = "SET NOCOUNT ON;"
'POSConnection.Execute LastSQL
'LastSQL = "SET LOCK_TIMEOUT 5000;"
'POSConnection.Execute LastSQL
'LastSQL = "SET XACT_ABORT ON;"
'POSConnection.Execute LastSQL
'
'LastSQL = "SET NOCOUNT ON;"
'Cn.Execute LastSQL
'LastSQL = "SET LOCK_TIMEOUT 5000;"
'Cn.Execute LastSQL
'LastSQL = "SET XACT_ABORT ON;"
'mLastStep = "Set XACT_ABORT ON for main connection"
'mLastSQL = LastSQL
'Cn.Execute LastSQL
'
''========================
'' Read options from central server
''========================
'Set rsDummy = New ADODB.Recordset
's = "SELECT * FROM TblOptions"
'LastSQL = s
'rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
'
'If Not rsDummy.EOF Then
'    NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
'    StoreDigit = Val(rsDummy!StoreDigit & "")
'    BranchDigit = Val(rsDummy!BranchDigit & "")
'    IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
'    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
'    InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
'    ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
'    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
'    NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
'    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
'    IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
'End If
'
'rsDummy.Close
'Set rsDummy = Nothing
'
'Exit Function
'
'ErrTrap:
'    On Error Resume Next
'
'    Extra = "ServerDb=" & ServerDb & vbCrLf & _
'            "POSDb=" & POSDb & vbCrLf & _
'            "DestinationServer=" & DestinationServer & vbCrLf & _
'            "SysSQLServerName=" & SysSQLServerName & vbCrLf & _
'            "POSServer=" & POSServer & vbCrLf & _
'            "SysSQLServerType=" & SysSQLServerType & vbCrLf & _
'            "SysSQLServerTypeTechnical=" & SysSQLServerTypeTechnical & vbCrLf & _
'            "Cn.ConnectionString=" & IIf(Cn Is Nothing, "", Cn.ConnectionString) & vbCrLf & _
'            "POS.ConnectionString=" & IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    LogErrDetailed "ConnectionFirst", Err, Erl, Extra, Cn, POSConnection, LastSQL
'
'    Text1 = IIf(Cn Is Nothing, "", Cn.ConnectionString)
'    Text2 = IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)
'
'    frmPopup.ShowMessage "ÕœÀ Œÿ√ ›Ì «·« ’«·"
'    ConnectionFirst = False
'End Function
Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean

    On Error GoTo ErrTrap

    Dim LastSQL As String
    Dim Extra As String
    Dim s As String
    Dim rsDummy As ADODB.Recordset

    ConnectionFirst = False
    LastSQL = ""

    '========================
    ' Read names first
    '========================
    ServerDb = Trim$(TxtServerDataBaseName.Text)
    DestinationServer = Trim$(SysSQLServerName)

    If Len(ServerDb) = 0 Then
        Err.Raise vbObjectError + 7001, , "«”„ ﬁ«⁄œ… »Ì«‰«  «·”Ì—›— «·„—ﬂ“Ì €Ì— „Õœœ"
    End If

    If Len(DestinationServer) = 0 Then
        Err.Raise vbObjectError + 7002, , "«”„ √Ê ⁄‰Ê«‰ «·”Ì—›— «·„—ﬂ“Ì €Ì— „Õœœ"
    End If

    '========================
    ' Close old main connection if open
    '========================
    On Error Resume Next
    If Not Cn Is Nothing Then
        If Cn.State <> adStateClosed Then Cn.Close
    End If
    Set Cn = Nothing
    On Error GoTo ErrTrap

    '========================
    ' Main DB Connection
    '========================
    Set Cn = New ADODB.Connection

    With Cn
        .CommandTimeout = 5000
        .CursorLocation = adUseServer
        .ConnectionTimeout = 5000

        If SysSQLServerType = 1 Then

            .ConnectionString = "Provider=SQLOLEDB.1;" & _
                                "Password=" & SysSQLServerUserpassword & ";" & _
                                "Persist Security Info=True;" & _
                                "User ID=" & SysSQLServerUserId & ";" & _
                                "Initial Catalog=" & ServerDb & ";" & _
                                "Data Source=" & DestinationServer

        ElseIf SysSQLServerType = 2 Then

            If SysSQLServerTypeTechnical = "0" Then
                .ConnectionString = "Provider=SQLOLEDB.1;" & _
                                    "Integrated Security=SSPI;" & _
                                    "Persist Security Info=False;" & _
                                    "Initial Catalog=" & ServerDb & ";" & _
                                    "Data Source=" & DestinationServer
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;" & _
                                    "Password=" & SysSQLServerUserpassword & ";" & _
                                    "Persist Security Info=True;" & _
                                    "User ID=" & SysSQLServerUserId & ";" & _
                                    "Initial Catalog=" & ServerDb & ";" & _
                                    "Data Source=" & DestinationServer
            End If

        Else
            Err.Raise vbObjectError + 7003, , "ﬁÌ„… SysSQLServerType €Ì— „⁄—Ê›…"
        End If

        .Errors.Clear
        .Open
    End With

    If IsLoad Then
        ConnectionFirst = True
        Exit Function
    End If

    '========================
    ' POS DB Connection
    '========================
    POSDb = Trim$(TxtPOSDB.Text)
    POSServer = Trim$(POSlServer.Text)

    If Len(POSDb) = 0 Then
        Err.Raise vbObjectError + 7004, , "«”„ ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ… €Ì— „Õœœ"
    End If

    If Len(POSServer) = 0 Then
        Err.Raise vbObjectError + 7005, , "«”„ ”Ì—›— «·‰ﬁÿ… €Ì— „Õœœ"
    End If

    '========================
    ' Close old POS connection if open
    '========================
    On Error Resume Next
    If Not POSConnection Is Nothing Then
        If POSConnection.State <> adStateClosed Then POSConnection.Close
    End If
    Set POSConnection = Nothing
    On Error GoTo ErrTrap

    Set POSConnection = New ADODB.Connection

    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000

        If SysSQLServerType = 1 Then

            .ConnectionString = "Provider=SQLOLEDB.1;" & _
                                "Password=" & SysSQLServerUserpassword & ";" & _
                                "Persist Security Info=True;" & _
                                "User ID=" & SysSQLServerUserId & ";" & _
                                "Initial Catalog=" & POSDb & ";" & _
                                "Data Source=" & POSServer

        ElseIf SysSQLServerType = 2 Then

            If SysSQLServerTypeTechnical = "0" Then
                .ConnectionString = "Provider=SQLOLEDB.1;" & _
                                    "Integrated Security=SSPI;" & _
                                    "Persist Security Info=False;" & _
                                    "Initial Catalog=" & POSDb & ";" & _
                                    "Data Source=" & POSServer
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;" & _
                                    "Password=" & SysSQLServerUserpassword & ";" & _
                                    "Persist Security Info=True;" & _
                                    "User ID=" & SysSQLServerUserId & ";" & _
                                    "Initial Catalog=" & POSDb & ";" & _
                                    "Data Source=" & POSServer
            End If

        Else
            Err.Raise vbObjectError + 7006, , "ﬁÌ„… SysSQLServerType €Ì— „⁄—Ê›… √À‰«¡ › Õ « ’«· «·‰ﬁÿ…"
        End If

        .Errors.Clear
        .Open
    End With

    '========================
    ' Session options
    '========================
    LastSQL = "SET NOCOUNT ON;"
    POSConnection.Execute LastSQL

    LastSQL = "SET LOCK_TIMEOUT 5000;"
    POSConnection.Execute LastSQL

    LastSQL = "SET XACT_ABORT ON;"
    POSConnection.Execute LastSQL

    LastSQL = "SET NOCOUNT ON;"
    Cn.Execute LastSQL

    LastSQL = "SET LOCK_TIMEOUT 5000;"
    Cn.Execute LastSQL

    LastSQL = "SET XACT_ABORT ON;"
    mLastStep = "Set XACT_ABORT ON for main connection"
    mLastSQL = LastSQL
    Cn.Execute LastSQL

    '========================
    ' Read options from central server
    '========================
    Set rsDummy = New ADODB.Recordset

    s = "SELECT * FROM TblOptions"
    LastSQL = s

    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rsDummy.EOF Then
        NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
        StoreDigit = Val(rsDummy!StoreDigit & "")
        BranchDigit = Val(rsDummy!BranchDigit & "")
        IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
        ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
        InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
        ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
        AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
        NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
        JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
        IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
    End If

    rsDummy.Close
    Set rsDummy = Nothing

    ConnectionFirst = True
    Exit Function

ErrTrap:
    On Error Resume Next

    Extra = "ServerDb=" & ServerDb & vbCrLf & _
            "POSDb=" & POSDb & vbCrLf & _
            "DestinationServer=" & DestinationServer & vbCrLf & _
            "SysSQLServerName=" & SysSQLServerName & vbCrLf & _
            "POSServer=" & POSServer & vbCrLf & _
            "SysSQLServerType=" & SysSQLServerType & vbCrLf & _
            "SysSQLServerTypeTechnical=" & SysSQLServerTypeTechnical & vbCrLf & _
            "Cn.ConnectionString=" & IIf(Cn Is Nothing, "", Cn.ConnectionString) & vbCrLf & _
            "POS.ConnectionString=" & IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)

    LogErrDetailed "ConnectionFirst", Err, Erl, Extra, Cn, POSConnection, LastSQL

    Text1 = IIf(Cn Is Nothing, "", Cn.ConnectionString)
    Text2 = IIf(POSConnection Is Nothing, "", POSConnection.ConnectionString)

    If Not rsDummy Is Nothing Then
        If rsDummy.State <> adStateClosed Then rsDummy.Close
        Set rsDummy = Nothing
    End If

    frmPopup.ShowMessage "ÕœÀ Œÿ√ ›Ì «·« ’«·"
    ConnectionFirst = False

End Function
Private Sub cmdTransfer_Click()
Command11_Click


Command2_Click

Command14_Click
POSname_Change
'Command12_Click
 'Command1_Click
 
cmdTransferMove_Click
' Command10_Click
 POSname_Change
End Sub

'

Private Sub cmdTransferMove2_Click()

    On Error GoTo ErrorHandler

    ' «· √ﬂœ „‰ ÊÃÊœ ‰ﬁÿ… „ ’·…
    If POSlServer.Text = "" Then
        frmPopup.ShowMessage "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«"
        Exit Sub
    End If
    If ConnectionFirst = False Then Exit Sub

    lblWait.Visible = True
    lblWait.Caption = "Ì „ «·«‰ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì…"
    DoEvents

    Dim POSConnection As New ADODB.Connection
   
    Dim rsTrans As New ADODB.Recordset
    Dim rsDetails As New ADODB.Recordset
    Dim rsValueAdded As New ADODB.Recordset
    Dim rsPayments As New ADODB.Recordset

    Dim BatchSize As Integer, recCounter As Integer
    BatchSize = 50
    recCounter = 0
Dim batchThreshold As Long, recCount As Long
    batchThreshold = 50
    recCount = 0

    Dim transBatchSQL As String, detailsBatchSQL As String, valueAddedBatchSQL As String, paymentsBatchSQL2 As String, paymentsBatchSQL As String, LastSQL As String
    transBatchSQL = ""
    detailsBatchSQL = ""
    valueAddedBatchSQL = ""

    Dim SessionCode As String
    SessionCode = Format(Now, "yyyymmddhhmmss")

    ' › Õ « ’«· POS
    POSConnection.CursorLocation = adUseServer
    POSConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
    POSConnection.Open

    ' › Õ « ’«· «·”Ì—›— «·„—ﬂ“Ì
    POSConnection.CursorLocation = adUseServer
'    Cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & mDBPOSName & ";Data Source=RemoteServer10"
'    Cn.Open

    ' › Õ Recordset ··„⁄«„·«  „‰ POS
    
    
   ' POSConnection.Execute "UPDATE Transactions SET SessionCode = '" & SessionCode & "' WHERE IsNull(Copied,0) =0 AND " & GetQuery

    
    'rsTrans.Open "SELECT * FROM Transactions WHERE IsNull(Copied,0) =0 AND SessionCode = '" & SessionCode & "' AND " & GetQuery & " ORDER BY Transaction_ID", POSConnection, adOpenForwardOnly, adLockReadOnly
    
        ' Ã·» «·»Ì«‰«  „‰ ÃœÊ· Transactions ›Ì «·”Ì—›—
    sql = "SELECT * FROM " & mServerD & "[Transactions] T2 " & _
          "WHERE Transaction_Type in (10,11) and NOT EXISTS (SELECT 1 FROM " & mPosD & "[Transactions] T1 WHERE T1.OldTransaction_ID = T2.Transaction_ID)"
    Set rs = New ADODB.Recordset
    'Cn.Open
    rs.Open sql, POSConnection, adOpenStatic, adLockReadOnly


    If rsTrans.EOF Then
        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
        GoTo EndSub
    End If


Dim CarOilChangeDate As Date, RecTime As Date, mTimeIn As String

    Dim mTimeStart As Date, mEndTime As Date, ActualDeliveryDate As Date, LatestDeliveryDate As Date
  '  mTimeStart = Now
  '  txtStartTime = mTimeStart
 '   Text3 = "Query: " & GetQuery
    POSConnection.Execute "SET XACT_ABORT ON;"
    POSConnection.BeginTrans

    Do While Not rsTrans.EOF

'        Dim currentDestTransactionID As Long
'        currentDestTransactionID = new_id("Transactions", "Transaction_ID", "", True) + recCounter
'
'        '  Ã„Ì⁄ Transaction
'        transBatchSQL = transBatchSQL & "INSERT INTO Transactions (Transaction_ID, Transaction_Date, Transaction_Type, PaymentType, CusID, StoreID, Emp_ID, BranchID, SessionCode) VALUES (" & _
'            currentDestTransactionID & "," & SQLDate(rsTrans("Transaction_Date"), True) & "," & rsTrans("Transaction_Type") & "," & rsTrans("PaymentType") & "," & rsTrans("CusID") & "," & rsTrans("StoreID") & "," & rsTrans("Emp_ID") & "," & rsTrans("BranchID") & ",'" & SessionCode & "');"
    recCount = recCount + 1
   ' ﬁ—«¡… «·ÕﬁÊ· „⁄ «·ﬁÌ„ «·«› —«÷Ì… ﬂ„« ›Ì «·ﬂÊœ «·√’·Ì
         Dim PayMentType As Long, cusID As Long, BranchID As Integer, BoxID As Long, BillBasedOn As Double
         Dim VAT As Double, VATYou As Double, NoteId As Long, Trans_DiscountType As Long
         Dim Trans_Discount As Double, TaxValue As Double, order_no As Long, SaleType As Long
         Dim TaxAddValue As Double, NetValue As Double, Transaction_NetValue As Double, DepandToConv As Long
         Dim CarTypeID As Long, OilsTypesID As Long, YearFact As Long, FixesAssetsID As Long, ColorID2 As Long
         Dim KM As Double, PPointID As Long, SupplerID As Long, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double
         Dim CarEnginoil As Double, CarGearOil As Double, InvoiceTypeCodeID As Long
         Dim storeID As Variant, userID As Variant, Emp_ID As Variant
         Dim NoteSerial As String, NoteSerial1 As String, TransactionComment As String
         Dim CashCustomerName As String, CashCustomerPhone As String
         Dim PlateNo As String, Shaseh As String, CarMeter As String
         Dim CIBAN As String
         Dim InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String
         Dim paymentnote As String, PaymentMeansCode As String
         Dim FromTransaction_Date As Date

         PayMentType = Val(rsTrans("PaymentType").Value & "")
          FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
         cusID = Val(rsTrans("CusID").Value & "")
         storeID = Val(rsTrans("StoreID").Value & "")
         userID = Val(rsTrans("UserID").Value & "")
         Emp_ID = Val(rsTrans("Emp_ID").Value & "")
         BranchID = Val(rsTrans("BranchID").Value & "")
         BoxID = Val(rsTrans("BoxID").Value & "")
         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
          PayMentType = Val(rsTrans("PaymentType").Value & "")
         cusID = Val(rsTrans("CusID").Value & "")
         storeID = Val(rsTrans("StoreID").Value & "")
         userID = Val(rsTrans("UserID").Value & "")
         Emp_ID = Val(rsTrans("Emp_ID").Value & "")
         BranchID = Val(rsTrans("BranchID").Value & "")
         BoxID = Val(rsTrans("BoxID").Value & "")
         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
         mTimeIn = Trim(rsTrans("TimeIn").Value & "")
         VAT = Val(rsTrans("VAT").Value & "")
         VATYou = Val(rsTrans("VATYou").Value & "")
         NoteSerial = rsTrans("NoteSerial").Value & ""
         NoteSerial1 = rsTrans("NoteSerial1").Value & ""
         NoteId = Val(rsTrans("NoteId").Value & "")
         FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
         TransactionComment = rsTrans("TransactionComment").Value & ""
         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
         FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
         TaxValue = Val(rsTrans("TaxValue").Value & "")
         order_no = Val(rsTrans("order_no").Value & "")
         SaleType = Val(rsTrans("SaleType").Value & "")
         CashCustomerName = rsTrans("CashCustomerName").Value & ""
         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
         NetValue = Val(rsTrans("NetValue").Value & "")
         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
         PlateNo = rsTrans("PlateNo").Value & ""
         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
         YearFact = Val(rsTrans("YearFact").Value & "")
         Shaseh = rsTrans("Shaseh").Value & ""
         CarMeter = rsTrans("CarMeter").Value & ""
         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
         ColorID2 = Val(rsTrans("ColorID2").Value & "")
         KM = Val(rsTrans("KM").Value & "")
         Chasee = rsTrans("Chasee").Value & ""
         PPointID = Val(rsTrans("PPointID").Value & "")
         Phone2 = rsTrans("Phone2").Value & ""
         SupplerID = Val(rsTrans("SupplerID").Value & "")
         Ser = Val(rsTrans("Ser").Value & "")
         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
         VAT = Val(rsTrans("VAT").Value & "")
         VATYou = Val(rsTrans("VATYou").Value & "")
         NoteSerial = rsTrans("NoteSerial").Value & ""
         NoteSerial1 = rsTrans("NoteSerial1").Value & ""
         NoteId = Val(rsTrans("NoteId").Value & "")
         TransactionComment = rsTrans("TransactionComment").Value & ""
         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
         TaxValue = Val(rsTrans("TaxValue").Value & "")
         order_no = Val(rsTrans("order_no").Value & "")
         SaleType = Val(rsTrans("SaleType").Value & "")
         CashCustomerName = rsTrans("CashCustomerName").Value & ""
         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
         NetValue = Val(rsTrans("NetValue").Value & "")
         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
         PlateNo = rsTrans("PlateNo").Value & ""
         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
         YearFact = Val(rsTrans("YearFact").Value & "")
         Shaseh = rsTrans("Shaseh").Value & ""
         CarMeter = rsTrans("CarMeter").Value & ""
         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
         ColorID2 = Val(rsTrans("ColorID2").Value & "")
         KM = Val(rsTrans("KM").Value & "")
         Chasee = rsTrans("Chasee").Value & ""
         PPointID = Val(rsTrans("PPointID").Value & "")
         Phone2 = rsTrans("Phone2").Value & ""
         SupplerID = Val(rsTrans("SupplerID").Value & "")
         Ser = Val(rsTrans("Ser").Value & "")
         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
         If Trim(rsTrans("CarOilChangeDate").Value & "") = "" Then
             CarOilChangeDate = Date
         Else
             CarOilChangeDate = rsTrans("CarOilChangeDate").Value & ""
         End If
         CIBAN = rsTrans("CIBAN").Value & ""
         'RecTime = IIf(rsTrans("RecTime").Value & "" = "", Time, rsTrans("RecTime").Value & "")
         Dim tmpRecTime As Variant
tmpRecTime = rsTrans("RecTime").Value
'If IsNull(tmpRecTime) Or Trim(CStr(tmpRecTime)) = "" Or tmpRecTime = "#12/30/1899#" Then
'    RecTime = Time
'Else
'    RecTime = tmpRecTime
'End If
'RecTime = IIf(IsNull(rsTrans("RecTime").Value), Now, rsTrans("RecTime").Value)

Dim tmpRecTimeStr As String
'tmpRecTimeStr = Trim(CStr(rsTrans("RecTime").Value & ""))
'
'If tmpRecTimeStr = "" Or tmpRecTimeStr = "30-Dec-1899" Then
'    RecTime = Time
'ElseIf IsDate(tmpRecTimeStr) Then
'    RecTime = CDate(tmpRecTimeStr)
'Else
'    RecTime = Time
'End If

Dim v As Variant: v = rsTrans("RecTime").Value
If IsDate(v) Then
    If Year(CDate(v)) = 1899 And Month(CDate(v)) = 12 And Day(CDate(v)) = 30 Then
        RecTime = Time
    Else
        RecTime = CDate(v)
    End If
Else
    RecTime = Time
End If


        ' RecTime = IIf(IsNull(rsTrans("RecTime").Value) Or Trim(rsTrans("RecTime").Value & "") = "", Time, rsTrans("RecTime").Value)
         ActualDeliveryDate = IIf(rsTrans("ActualDeliveryDate").Value & "" = "", Date, rsTrans("ActualDeliveryDate").Value & "")
         LatestDeliveryDate = IIf(rsTrans("LatestDeliveryDate").Value & "" = "", Date, rsTrans("ActualDeliveryDate").Value & "")
         
         InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
         InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
         DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
         TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
         paymentnote = rsTrans("paymentnote").Value & ""
         PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""
         FromTransaction_Date = rsTrans("Transaction_Date").Value

         ' ≈–« POSBillType = 0°  ⁄œÌ· NoteSerial ÊNoteId
         If Val(rsTrans("POSBillType").Value & "") = 0 Then
             NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
             NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
         End If

         TransactionComment = " ”‰œ  ÕÊÌ· „Œ“‰Ì „‰ﬁÊ·… „‰ " & "«·”Ì—›—" & "   " & _
                              "   —ﬁ„ «·›« Ê—… " & NoteSerial1
   '  Ê·Ìœ —ﬁ„ ÃœÌœ ··„⁄«„·… ⁄·Ï «·ÊÃÂ…
'         Dim currentDestTransactionID As String
'         currentDestTransactionID = CStr((new_id("Transactions", "Transaction_ID", "", True) + recCount))
        Dim currentDestTransactionID As String
        Dim rsSer As ADODB.Recordset
        
        ' ‰œ«¡ «·” Ê—œ »—Ê”ÌÃ—
        Set rsSer = POSConnection.Execute("EXEC dbo.ReserveTransactionId")
        
        If Not (rsSer.EOF) Then
            currentDestTransactionID = CStr(rsSer.Fields("NewId").Value)
        Else
            Err.Raise vbObjectError + 500, , "·„ Ì „ ≈—Ã«⁄ Transaction_ID ÃœÌœ „‰ «·”Ì—›—"
        End If
        
        rsSer.Close
        Set rsSer = Nothing


transSQL = "INSERT INTO " & POSConnection & "Transactions (" & _
"Transaction_ID, Transaction_Date,TimeIn ,TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
"CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, NoteSerial1, " & _
"NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, OldNoteserial, OldNoteId, " & _
"OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, SaleType, CashCustomerName, " & _
"TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, DepandToConv, CarTypeID, " & _
"PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, ColorID2, KM, Chasee, PPointID, Phone2, " & _
"SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, CarGearOil, CarOilChangeDate, CIBAN, RecTime, " & _
"ActualDeliveryDate, LatestDeliveryDate, InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, " & _
"TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("

transSQL = transSQL & currentDestTransactionID & "," & SQLDate(FromTransaction_Date, True) & ",'" & Trim(mTimeIn) & "'," & _
Val(rsTrans("TypeInvoice") & "") & ",'" & Replace(rsTrans("Transaction_Serial") & "", "'", "''") & "'," & _
FromTransaction_Type & "," & PayMentType & "," & cusID & "," & storeID & "," & userID & "," & _
Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & ",'" & _
NoteSerial & "','" & NoteSerial1 & "'," & NoteId & ",1,'" & Replace(TransactionComment, "'", "''") & "','" & _
SessionCode & "'," & IIf(Val(rsTrans("POSBillType") & "") = 0, 1, Val(rsTrans("POSBillType") & "")) & ",'" & rsTrans("Noteserial1") & "" & "','" & Trim(rsTrans("Noteserial") & "") & "'," & _
Val(rsTrans("NoteId") & "") & "," & rsTrans("Transaction_ID") & "," & Trans_DiscountType & "," & Val(Trans_Discount & "") & "," & _
TaxValue & ",'" & order_no & "'," & SaleType & ",'" & Replace(cleanCashCustomerName, "'", "''") & "'," & _
TaxAddValue & ",'" & CashCustomerPhone & "'," & SQLDate(rsTrans("last_changed"), True) & ","

transSQL = transSQL & NetValue & "," & Transaction_NetValue & "," & IIf(DepandToConv, 1, 0) & "," & _
CarTypeID & ",'" & PlateNo & "'," & OilsTypesID & "," & YearFact & ",'" & Shaseh & "','" & CarMeter & "'," & _
FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Chasee & "'," & PPointID & ",'" & Phone2 & "'," & _
SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & CarEnginoil & "," & _
CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & CIBAN & "'," & SQLDate(RecTime, True) & ","

transSQL = transSQL & SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
InvoiceTypeCodeID & ",'" & InvoiceTypeCodename & "','" & DocumentCurrencyCode & "','" & _
TaxCurrencyCode & "','" & Replace(paymentnote, "'", "''") & "','" & PaymentMeansCode & "')"

transBatchSQL = transBatchSQL & transSQL & vbCrLf

        '  ›«’Ì· Transaction
'        rsDetails.Open "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & rsTrans("Transaction_ID"), POSConnection
'        Do While Not rsDetails.EOF
'            detailsBatchSQL = detailsBatchSQL & "INSERT INTO Transaction_Details (Transaction_ID, Item_ID, Quantity, Price, SessionCode) VALUES (" & _
'                currentDestTransactionID & "," & rsDetails("Item_ID") & "," & rsDetails("Quantity") & "," & rsDetails("Price") & ",'" & SessionCode & "');"
'            rsDetails.MoveNext
'        Loop
'        rsDetails.Close

'Dim rsDetails As New ADODB.Recordset
sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID"))
rsDetails.CursorType = adOpenForwardOnly
rsDetails.LockType = adLockReadOnly
rsDetails.Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

Do While Not rsDetails.EOF
    Dim detailSQL As String

    detailSQL = "INSERT INTO " & POSConnection & "Transaction_Details (" & _
        "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
        "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
        "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES ("

    detailSQL = detailSQL & currentDestTransactionID & "," & Val(rsDetails("Item_ID")) & "," & Val(rsDetails("ItemCase") & "") & "," & _
        Val(rsDetails("Quantity") & "") & "," & Val(rsDetails("Price") & "") & "," & Val(rsDetails("ItemDiscountType") & "") & "," & _
        Val(rsDetails("ItemDiscount") & "") & "," & Val(rsDetails("ShowQty") & "") & "," & Val(rsDetails("showPrice") & "") & "," & _
        Val(rsDetails("UnitId")) & "," & Val(rsDetails("ColorID") & "") & "," & Val(rsDetails("ItemSize") & "") & "," & _
        Val(rsDetails("ClassId") & "") & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo") & "") & "," & _
        Val(rsDetails("PumpId") & "") & "," & Val(rsDetails("PrevQty") & "") & ",'" & Replace(Trim(rsDetails("PrintName") & ""), "'", "''") & "'," & _
        Val(rsDetails("Cash") & "") & "," & Val(rsDetails("Mada") & "") & "," & Val(rsDetails("Visa") & "") & "," & _
        Val(rsDetails("Deferred") & "") & "," & Val(rsDetails("AmountH") & "") & "," & Val(rsDetails("AmountHComm") & "") & ","

    detailSQL = detailSQL & "'" & Replace(Trim(rsDetails("DetailsPump") & ""), "'", "''") & "','" & Replace(Trim(rsDetails("Account_CodeComm") & ""), "'", "''") & "','" & _
        Replace(Trim(rsDetails("Account_Code") & ""), "'", "''") & "'," & IIf(rsDetails("IsOther").Value, 1, 0) & ")"

    detailsBatchSQL = detailsBatchSQL & detailSQL & vbCrLf

    rsDetails.MoveNext
Loop

rsDetails.Close



'        ' ﬁÌ„… „÷«›…
'        rsValueAdded.Open "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & rsTrans("Transaction_ID"), POSConnection
'        Do While Not rsValueAdded.EOF
'            valueAddedBatchSQL = valueAddedBatchSQL & "INSERT INTO TransactionValueAdded (Transaction_ID, ItemID, VAT, Valu, SessionCode) VALUES (" & _
'                currentDestTransactionID & "," & rsValueAdded("ItemID") & "," & rsValueAdded("VAT") & "," & rsValueAdded("Valu") & ",'" & SessionCode & "');"
'            rsValueAdded.MoveNext
'        Loop
'        rsValueAdded.Close






        recCounter = recCounter + 1
    If recCounter Mod BatchSize = 0 Then
        If transBatchSQL <> "" Then
            LastSQL = transBatchSQL
            WriteLog "Executing transBatchSQL batch", transBatchSQL
            POSConnection.Execute transBatchSQL
        End If
    
        If detailsBatchSQL <> "" Then
            LastSQL = detailsBatchSQL
            WriteLog "Executing detailsBatchSQL batch", detailsBatchSQL
            POSConnection.Execute detailsBatchSQL
        End If
    
     
    
        transBatchSQL = "": detailsBatchSQL = "": valueAddedBatchSQL = "": paymentsBatchSQL2 = "": paymentsBatchSQL = ""
    End If


        rsTrans.MoveNext
    Loop

If transBatchSQL <> "" Then
    LastSQL = transBatchSQL
    WriteLog "Executing transBatchSQL final", transBatchSQL
    POSConnection.Execute transBatchSQL
End If

If detailsBatchSQL <> "" Then
    LastSQL = detailsBatchSQL
    WriteLog "Executing detailsBatchSQL final", detailsBatchSQL
    POSConnection.Execute detailsBatchSQL
End If



transBatchSQL = "": detailsBatchSQL = "": valueAddedBatchSQL = "": paymentsBatchSQL2 = "": paymentsBatchSQL = ""

POSConnection.CommitTrans

'=== [2] POS source counters for this Session ===
Dim SrcHeads As Long, SrcDet As Long, SrcVAT As Long, SrcPay As Long, SrcPay2 As Long
Dim rsCnt As ADODB.Recordset


Dim elapsedSec As Long, elapsedMin As Long

elapsedSec = DateDiff("s", mTimeStart, Now)
elapsedMin = elapsedSec \ 60
elapsedSec = elapsedSec Mod 60
frmPopup.ShowMessage " „ «·‰ﬁ· »‰Ã«Õ." & vbCrLf & _
       "«·Êﬁ  «·„” €—ﬁ: " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì….", vbInformation

lblWait.Caption = " „ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì… »‰Ã«Õ." & _
                  " «·Êﬁ  «·„” €—ﬁ: " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì…."
txtEndTime = CStr(Now)
    POSname_Change
   ' MsgBox " „ «·‰ﬁ· »‰Ã«Õ"

EndSub:
    'lblWait.Visible = False
    Exit Sub

ErrorHandler:
    Cn.RollbackTrans
    POSConnection.Execute "UPDATE Transactions SET Copied = null, SessionCode = null WHERE SessionCode = '" & SessionCode & "'"
    WriteLog "ErrorHandler: " & Err.Description, LastSQL
    frmPopup.ShowMessage "Œÿ√ «À‰«¡ «·‰ﬁ· —Ã«¡ «· Ê«’· „⁄ „”∆Ê·Ï «·‰Ÿ«„: " & Err.Description, vbCritical
    lblWait.Visible = False

End Sub

'
'
'
'
Private Sub cmdUdateFiles_Click()



On Error GoTo EE:
'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
frmPopup.ShowMessage "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   
  UpdateFiles POSlServer, POSDb, "cachierData", "Id"
  UpdateFiles POSlServer, POSDb, "TblStore", "StoreID"
  UpdateFiles POSlServer, POSDb, "TblBoxesData", "BoxId"
  UpdateFiles POSlServer, POSDb, "BanksData", "BankId"
  UpdateFiles POSlServer, POSDb, "TblEmployee", "Emp_ID"
  
'  UpdateFiles POSlServer, POSDb, "ACCOUNTS", "Account_ID"
  UpdateFiles POSlServer, POSDb, "TblCustemers", "CusId"
  
   
 UpdateFiles POSlServer, POSDb, "TblUsers", "UserID"
  UpdateFiles POSlServer, POSDb, "TblBranchesData", "branch_id"
  
  UpdateFiles POSlServer, POSDb, "TblUsersBoxes", "id"
  UpdateFiles POSlServer, POSDb, "TblUsersBranches", "id"
  UpdateFiles POSlServer, POSDb, "TblUserScreen", "id"
  UpdateFiles POSlServer, POSDb, "TblUsersStores", "id"
  
'UpdateFiles POSlServer, POSDb, "TblOptions", "PlayNotesAlramSound"
  
  UpdateFiles POSlServer, POSDb, "TblLink_Item_To_StoreH", "Ind"
  UpdateFiles POSlServer, POSDb, "TblLink_Item_To_Store_Details1", "id"
  UpdateFiles POSlServer, POSDb, "TblLink_Item_To_Store_Details2", "id2"
  UpdateFiles POSlServer, POSDb, "TblLink_Item_To_Store_Details3", "id"
  
  
  

  
  If LastIdentityInsertTable <> "" Then
    POSConnection.Execute "SET IDENTITY_INSERT dbo." & LastIdentityInsertTable & " OFF"
    LastIdentityInsertTable = ""
End If

  frmPopup.ShowMessage " „ ‰ﬁ· «·»Ì«‰«  «·«”«”Ì…"
   Exit Sub
EE:
frmPopup.ShowMessage "BasicData"
End Sub


Private Sub cmdUpdatePrice_Click()
Dim StrSQL As String
If POSlServer.Text = "" Then
frmPopup.ShowMessage "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   'step one check item
       
    ss = "     USE " & ServerDb & vbNewLine
     mLastStep = "Insert missing items into remote server"
mLastSQL = ss
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    
             'checkGroup
        Dim NoOfGroups_pos As Double
        Dim NoOfGroups_server As Double
             
        Dim MaxGroupid_pos As Double
        Dim MaxGroupidserver As Double
        
                       
     
        
         'MsgBox "Step 3"
        Dim s As String
        
    
            
   
            BolFrmLoaded = True
    
            
      POSConnection.CommandTimeout = 1000
     Cn.CommandTimeout = 1000
              
            s = ""
             
       
       
       
       
       
      Dim mPosD As String
Dim mServerD As String
mPosD = "[" & POSlServer & "]." & POSDb & ".dbo."
mServerD = "[" & SysSQLServerName & "]." & ServerDb & ".dbo."

' ??? ???????? ??????? - TblUnites
s = "INSERT INTO " & mPosD & "TblUnites (columns...) " & _
    "SELECT columns... " & _
    "FROM " & mServerD & "TblUnites T2 " & _
    "LEFT JOIN " & mPosD & "TblUnites T1 ON T2.UnitID = T1.UnitID " & _
    "WHERE T1.UnitID IS NULL;"
    
     mLastStep = "Insert missing items into remote server"
mLastSQL = s
Cn.Execute s

' ??? ???????? ??????? - TblItemsUnits
s = "INSERT INTO " & mPosD & "TblItemsUnits (columns...) " & _
    "SELECT columns... " & _
    "FROM " & mServerD & "TblItemsUnits T2 " & _
    "LEFT JOIN " & mPosD & "TblItemsUnits T1 ON T2.ItemID = T1.ItemID " & _
    "WHERE T1.ItemID IS NULL;"
        
     mLastStep = "Insert missing items into remote server"
mLastSQL = s
Cn.Execute s

' ??? ???????? ??????? - TblItems
s = "INSERT INTO " & mPosD & "TblItems (ItemID, ItemName, barCodeNO, Code, Fullcode, IsArchive) " & _
    "SELECT T2.ItemID, T2.ItemName, T2.barCodeNO, T2.Code, T2.Fullcode, ISNULL(T2.IsArchive, 0) " & _
    "FROM " & mServerD & "TblItems T2 " & _
    "LEFT JOIN " & mPosD & "TblItems T1 ON T2.ItemID = T1.ItemID " & _
    "WHERE T1.ItemID IS NULL;"
        
     mLastStep = "Insert missing items into remote server"
mLastSQL = s
Cn.Execute s

' ????? ???????? - TblItemsUnits
s = "UPDATE T1 " & _
    "SET T1.UnitSalesPrice = T2.UnitSalesPrice, " & _
    "    T1.MaxSelingPrice = T2.MaxSelingPrice, " & _
    "    T1.UnitWholeSalePrice = T2.UnitWholeSalePrice, " & _
    "    T1.MinSelingPrice = T2.MinSelingPrice, " & _
    "    T1.UnitPurPrice = T2.UnitPurPrice " & _
    "FROM " & mPosD & "TblItemsUnits T1 " & _
    "INNER JOIN " & mServerD & "TblItemsUnits T2 " & _
    "ON T1.ItemID = T2.ItemID AND T1.UnitId = T2.UnitId;"
        
     mLastStep = "Insert missing items into remote server"
mLastSQL = s
Cn.Execute s

' ????? ???????? - TblItems
s = "UPDATE T1 " & _
    "SET T1.ItemName = T2.ItemName, " & _
    "    T1.barCodeNO = T2.barCodeNO, " & _
    "    T1.Code = T2.Code, " & _
    "    T1.Fullcode = T2.Fullcode, " & _
    "    T1.IsArchive = ISNULL(T2.IsArchive, 0) " & _
    "FROM " & mPosD & "TblItems T1 " & _
    "INNER JOIN " & mServerD & "TblItems T2 " & _
    "ON T1.ItemID = T2.ItemID;"
        
     mLastStep = "Insert missing items into remote server"
mLastSQL = s
Cn.Execute s

frmPopup.ShowMessage "?? ??? ?????? ?????? ??????? ?????"
 
       
       
       
       Exit Sub
            
            
            
            
 
            'POSConnection.Execute "Delete " & mPosD & "TblItemsUnits "
 
           ' MsgBox "Step 4"
            s = " INSERT INTO " & mPosD & "TblUnites"
            s = s & " SELECT *"
            s = s & " FROM   " & mServerD & "TblUnites T2"
            s = s & " WHERE  T2.UnitID NOT IN (SELECT UnitID"
            s = s & "                                      FROM   " & mPosD & "TblUnites);"
            
            Cn.Execute s
         
'
            s = " INSERT INTO " & mPosD & "TblItemsUnits"
            s = s & " SELECT *"
            s = s & " FROM   " & mServerD & "TblItemsUnits T2"
            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
            s = s & "                                      FROM   " & mPosD & "TblItemsUnits);"
                                     
            Cn.Execute s
            
            
             

             
'
'        sql = " select * from TblItemsUnits    "
'        Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly
'        Do While Not Rs3.EOF
'
'             s = " Update " & mPosD & "TblItemsUnits Set UnitSalesPrice = " & Val(Rs3!UnitSalesPrice & "")
'             s = s & " ,MaxSelingPrice = " & Val(Rs3!MaxSelingPrice & "")
'             s = s & " ,UnitWholeSalePrice= " & Val(Rs3!UnitWholeSalePrice & "")
'             s = s & " ,MinSelingPrice = " & Val(Rs3!MinSelingPrice & "")
'             s = s & " ,UnitPurPrice= " & Val(Rs3!UnitPurPrice & "")
'
'
'             s = s & " Where TblItemsUnits.ItemID = " & Val(Rs3!ItemID & "")
'             s = s & " And TblItemsUnits.UnitId = " & Val(Rs3!UnitId & "")
'             Cn.Execute s
'            Rs3.MoveNext
'        Loop
        
    

    
    
                  
             frmPopup.ShowMessage " „ ‰ﬁ· »Ì«‰«  «·ÊÕœ« "
         
    
            
       ' Rs3.Close
        
        sql = " select IsNull(IsArchive,0) IsArchive2,ItemName,ItemID,barCodeNO,Code,Fullcode from TblItems    "
        sql = sql & " where ItemId In (Select  FF.ItemId from " & mPosD & "TblItems FF where ff.ItemName <> TblItems.ItemName "
        sql = sql & " Or ff.barCodeNO <> TblItems.barCodeNO Or ff.Code <> TblItems.Code) "
        Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly
        Do While Not Rs3.EOF
             
             s = " Update " & mPosD & "TblItems Set IsArchive = " & IIf(Rs3!IsArchive2, 1, 0)
             s = s & " ,ItemName = '" & Trim(Rs3!ItemName & "") & "'"
             s = s & " ,barCodeNO= '" & Trim(Rs3!barCodeNO & "") & "'"
              s = s & " ,barCodeNO= '" & Trim(Rs3!barCodeNO & "") & "'"
              s = s & " ,Code= '" & Trim(Rs3!Code & "") & "'"
             
             
             s = s & " Where TblItems.ItemID = " & Rs3!ItemID
            
             
             Cn.Execute s
            Rs3.MoveNext
        Loop
            
        
            
            

            
            
             frmPopup.ShowMessage " „ ‰ﬁ· »Ì«‰«  «·«’‰«›"
             Command2.Enabled = False
    
        
  
    
                
            
            
         
            
End Sub

Private Sub cmdUpdateSerial_Click()
Dim s As String
Dim rsDummy As New ADODB.Recordset
Dim rsDummyMax As New ADODB.Recordset
Dim TransType  As Long
Dim mYear As Integer
Dim Month As Integer
Dim BranchID As Integer
TransType = IIf(optSales, 21, 22)
BranchID = Val(txtBranch)
Month = Val(txtMonth)
mYear = Val(txtYear)

    Dim mPosD As String
    Dim mServerD As String
     mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
     mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
     mServerD = ServerDb & ".dbo."
     
Dim mLastSerial As Double

Dim mLastSerialString As String
Dim intIal As String
Dim InBranch As String
Dim InUser As String
Dim InitMonth As String
Dim InitUser As String

If BranchID < 10 Then
    InBranch = "0" + CStr(BranchID)
  

Else
    
    InBranch = CStr(BranchID)
End If

If Month < 10 Then

    InitMonth = "0" + CStr(Month)
Else
    InitMonth = CStr(Month)
End If





        
s = " SELECT MAX(NoteSerial1) as aa"
s = s & " From Transactions"
s = s & " WHERE  Transaction_Type = " & TransType
s = s & "        AND MONTH(Transaction_Date) = " & Month
s = s & "                    AND YEAR(Transaction_Date) =" & mYear
s = s & "                    AND BranchId = " & BranchID
's = s & "                    AND SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50)), 1, 1) = " & BranchID
s = s & "                    AND ISNULL(NoteSerial1, '0') <> '0'"
s = s & "                    AND ISNULL(NoteSerial1, '') <> ''"
's = s & "                     AND BranchId = 9898"


s = s & "                     "
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, POSConnection, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    mLastSerial = Val(rsDummy!aa & "") + 1
    If mLastSerial = 0 Then
        mLastSerial = CDbl(intIal)
    End If
Else
    If mLastSerial = 0 Then
        mLastSerial = CLng(intIal)
    End If
End If
mLastSerial = 0
intIal = "1" & InUser & InBranch + "20" + InitMonth + "001"
mLastSerialString = CStr(mLastSerial)

rsDummy.Close

    
    s = " SELECT *"
    s = s & " From Transactions"
    s = s & " WHERE  Transaction_Type = " & TransType
    s = s & "            AND YEAR(Transaction_Date) = " & mYear
    s = s & "            AND MONTH(Transactions.Transaction_Date) = " & Month
    s = s & "            AND BranchId = " & BranchID
    s = s & "            AND ISNULL(NoteSerial1, '0') IN (SELECT ISNULL(t.NoteSerial1, '0')"
    s = s & "                                             FROM   Transactions AS t"
    s = s & "                                             WHERE  t.Transaction_Type = " & TransType
    s = s & "                                                    AND YEAR(t.Transaction_Date) = " & mYear
    s = s & "                                                    AND MONTH(Transaction_Date) = " & Month
    s = s & "                                                    AND ISNULL(t.NoteSerial1, '0') IN (SELECT ISNULL(d.NoteSerial1, '0')"
    s = s & "                                                                                       FROM   Transactions d"
    s = s & "                                                                                       WHERE  d.Transaction_Type ="
    s = s & "                                                                                              " & TransType
    s = s & "                                                                                              AND MONTH(Transaction_Date) ="
    s = s & "                                                                                                  " & Month
    s = s & "                                                                                              AND BranchId = " & BranchID
    s = s & "                                                                                              AND YEAR(d.Transaction_Date) ="
    s = s & "                                                                                                  " & mYear & " )"
    s = s & "                                                    AND BranchId = " & BranchID
    s = s & "                                             Group By"
    s = s & "                                                    NoteSerial1"
    s = s & "                                             HAVING (COUNT(*) > 1))"




    rsDummy.Open s, POSConnection, adOpenKeyset, adLockOptimistic
    
    If rsDummy.EOF Then
        rsDummy.Close
        s = "Select *"
        s = s & " From Transactions"
        s = s & " WHERE  Transaction_Type = " & TransType
        s = s & "            AND YEAR(Transaction_Date) = " & mYear
        s = s & "            AND MONTH(Transactions.Transaction_Date) = " & Month
        s = s & "            AND BranchId = " & BranchID
        s = s & "            AND (IsNull(NoteSerial1,'') = '' Or IsNull(NoteSerial1,'0') = '0')"
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockOptimistic
    End If
    Do While Not rsDummy.EOF
    
        If Val(rsDummy!userID & "") < 10 Then
            InUser = "0" + CStr(Val(rsDummy!userID & ""))
        
        
        Else
        
            InUser = CStr(Val(rsDummy!userID & ""))
        End If


             
        s = " SELECT Max(CAST (NoteSerial1 AS BIGINT )) as aa"
        s = s & " From Transactions"
        s = s & " WHERE  Transaction_Type = " & TransType
        s = s & "        AND MONTH(Transaction_Date) = " & Month
        s = s & "                    AND YEAR(Transaction_Date) =" & mYear
        s = s & "                    AND BranchId = " & BranchID
        s = s & "                    AND UserId = " & Val(InUser)
        's = s & "                    AND SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50)), 1, 1) = " & BranchID
        s = s & "                    AND ISNULL(NoteSerial1, '0') <> '0'"
        s = s & "                    AND ISNULL(NoteSerial1, '') <> ''"
        's = s & "                     AND BranchId = 9898"
        
        
        s = s & "              and     NoteSerial1  <> "
          s = s & " (SELECT Max(CAST (NoteSerial1 AS BIGINT )) +1 as aa"
        s = s & " From Transactions"
        s = s & " WHERE  Transaction_Type = " & TransType
        s = s & "        AND MONTH(Transaction_Date) = " & Month
        s = s & "                    AND YEAR(Transaction_Date) =" & mYear
        s = s & "                    AND BranchId = " & BranchID
        s = s & "                    AND UserId = " & Val(InUser)
        's = s & "                    AND SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50)), 1, 1) = " & BranchID
        s = s & "                    AND ISNULL(NoteSerial1, '0') <> '0'"
        s = s & "                    AND ISNULL(NoteSerial1, '') <> '')"
        's = s & "                     AND BranchId = 9898"
        
        Set rsDummyMax = New ADODB.Recordset
        rsDummyMax.Open s, POSConnection, adOpenStatic, adLockReadOnly
        If Not rsDummyMax.EOF Then
            mLastSerial = Val(rsDummyMax!aa & "") + 1
            If mLastSerial = 0 Then
                mLastSerial = CDbl(intIal)
            End If
        Else
            If mLastSerial = 0 Then
                mLastSerial = CDbl(intIal)
            End If
        End If
        mLastSerialString = CStr(mLastSerial)

         rsDummy!NoteSerial1 = mLastSerial
         rsDummy!TransactionComment = "Test2"
         rsDummy.Update
         If Val(rsDummy!NoteId & "") <> 0 Then
            s = "UPDATE Notes Set NoteSerial1 = " & mLastSerial & " Where NoteId = " & Val(rsDummy!NoteId)
                
     mLastStep = "Insert missing items into remote server"
mLastSQL = s
            POSConnection.Execute s
        End If
    
        If Val(rsDummy!Copied & "") <> 0 And TransType = 21 Then
            s = " Update   " & mServerD & "Transactions Set OldNoteSerial1 = " & mLastSerial
            s = s & " Where OldTransaction_ID = " & Val(rsDummy!Transaction_ID & "")
            s = s & "  AND BranchId = " & BranchID
            s = s & " and Transaction_Type = " & TransType
            Cn.Execute s
        End If
        mLastSerial = mLastSerial + 1
        
        rsDummy.MoveNext
    Loop

   frmPopup.ShowMessage " „ ÷»ÿ «·”Ì—Ì«·"
 

End Sub

Private Sub Command1_Click()
On Error GoTo ErrTrap
Dim StrSQL As String
'On Error GoTo ErrTrap
If POSlServer.Text = "" Then
    frmPopup.ShowMessage "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«"
    Exit Sub
End If


If ConnectionFirst = False Then
Exit Sub
End If
Dim X As Date
Dim mTimeStart As String
'ServerDb = DestinationServer
 'POSDb = TxtServerDataBaseName.Text
    lblWait.Visible = True
  ' Command2_Click
    Dim rs As New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    JLCodeBasedOnBranch = IIf(rs("JLCodeBasedOnBranch").Value = 0 Or IsNull(rs("JLCodeBasedOnBranch").Value), False, True)
    StoreDigit = IIf(IsNull(rs("StoreDigit").Value), 1, (rs("StoreDigit").Value))
    BranchDigit = IIf(IsNull(rs("BranchDigit").Value), 1, (rs("BranchDigit").Value))
    

    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If SysSQLServerType = 1 Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'            "Persist Security Info=False;Initial Catalog=" & POSDb & _
'            ";Data Source=" & POSlServer & ";Port=1433"
'
        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
        ElseIf SysSQLServerType = 2 Then
             If SysSQLServerTypeTechnical = "0" Then
             .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                "Persist Security Info=False;Initial Catalog=" & POSDb & _
                ";Data Source=" & POSlServer & ";Port=1433"
              Else
                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer 'SysSQLServerName
            End If
        End If
       '   Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Adnan;Data Source=WAELPC\SQLEXPRESS;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=WAELPC;Use Encryption for Data=False;Tag with column collation when possible=False;

        .Open
    End With


GoTo Transactions

Transactions:
Dim SessionCode As String
Dim mMaxNo As Long
Dim ss As String
Dim rsDummyMax As New ADODB.Recordset
 Dim BeginTrans As Boolean
Dim isFoundData As Boolean

'ss = "Select Max(SessionCode ) MaxNo from TblOffline"
'rsDummyMax.Open ss, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'If rsDummyMax.EOF Then
'    mMaxNo = Val(rsDummyMax!MaxNo & "") + 1
    
'End If

SessionCode = CStr(Now) '& mMaxNo


'////////////////////////////////////////copy Sales Transactions

    Dim Rs3 As ADODB.Recordset
    Dim rsDouble_Entry As ADODB.Recordset
    
    Set Rs3 = New ADODB.Recordset
    
    
    Dim sql As String
    Dim mytext As String
    
   
   ' sql = " select * from Transactions    WHERE  Copied is null And " & GetQuery
    sql = " select * from Transactions    WHERE  Copied is null And POSBillType = 1 and " & GetQuery
    sql = " select * from Transactions    WHERE   Copied is null  And " & GetQuery
    
'    Dim tempString As String
'    Dim i As Integer
'    tempString = "0"
'    For i = 0 To Me.SelectedTransTypeList.ListCount - 1
'        tempString = tempString & "," & Me.SelectedTransTypeList.ItemData(i)
'    Next i
'    GetTransIds = tempString
    
    
    
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    mTimeStart = Now
    txtStartTime = mTimeStart
    Text3 = sql
    Dim FromTransaction_ID As Double
    Dim FromBranchID As Integer
    Dim FromTransaction_Date As Date
    Dim last_changed As Date
    
    Dim FromNots As String
    Dim FromNots2 As String
    Dim fromTransaction_Serial As String
    Dim FromNoteseial1 As String
    Dim FromTransaction_Type As Integer
    
    Dim BranchID As Integer
    Dim Transaction_ID As Double
    Dim Transaction_Type As Integer
    Dim Transaction_Date As Date
    Dim Transaction_Serial  As String
    Dim Nots As String
    Dim Nots2 As String
    Dim mTransaction_NetValue As Double
    Dim DepandToConv As Boolean
    Dim TypeInvoice As Integer
'eee
    'Dim Transaction_Type As Integer
    Dim FromNoteId As Double
   
    
   
 'sales
    If chkSales.Value = vbChecked Or chkSalesReturn.Value = vbChecked Or chkPurchase.Value = vbChecked Or chkPurchaseReturn.Value = vbChecked Or chkSalesOffers.Value = vbChecked Then
        If Rs3.RecordCount > 0 Then
'            Set cProgress = New ClsProgress
'            BolFrmLoaded = True
'            cProgress.ProgressType = Waiting
'            cProgress.StartProgress

'            Do While Rs3.State = adStateExecuting
'                DoEvents
'            Loop
            
'            If BolFrmLoaded = True Then
'                cProgress.StopProgess
'                Set cProgress = Nothing
'            End If

Dim rsCus As New ADODB.Recordset
Dim rsCus2 As New ADODB.Recordset

                Cn.BeginTrans
                BeginTrans = True
                
               ' MsgBox Rs3.RecordCount
                
                For i = 1 To Rs3.RecordCount
                    
                    
                    FromTransaction_Type = IIf(IsNull(Rs3("Transaction_Type").Value), 0, Rs3("Transaction_Type").Value)
                    FromTransaction_ID = IIf(IsNull(Rs3("Transaction_ID").Value), 0, Rs3("Transaction_ID").Value)
                    
                    Dim issueNoteid As String
                    Dim issuenoteserial As String
                    Dim issuenoteserial1 As String
                    Dim FromEmp_ID As Double
                    
                    Dim FromStoreID As Double
                    Dim FromCusID As Double
                    Dim FromBoxid As Double
                    Dim PayMentType As Integer
                    Dim BillBasedOn
                    'Dim BillBasedOn As Integer
                    Dim VATYou As Double
                    Dim VAT As Double
                    Dim FromUserID As Double
                    Dim POSBillType As Double
                    
                    Dim Trans_DiscountType As Integer
                    Dim Trans_Discount As Double
                    Dim TaxValue As Double
                    Dim order_no As String
                    Dim SaleType As Integer
                    Dim CashCustomerName As String
                    Dim TaxAddValue As Double
                    Dim CashCustomerPhone As String
                    Dim NetValue As Double
                    
                    Dim CarTypeID As Long, PlateNo As String, OilsTypesID As Long, YearFact As Long, Shaseh As String, CarMeter As String, FixesAssetsID As Long, ColorID2 As Integer, KM As Double, Chasee As String, PPointID As Long _
                    , Phone2 As String, SupplerID As Integer, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double, CarEnginoil As Double, CarGearOil As Double, CarOilChangeDate As Date
                    
                    
                    
Dim PumpId As Long, PrevQty As Double, PrintName As String, Cash As Double, Mada As Double, Visa As Double, Deferred As Double, AmountH As Double, AmountHComm As Double, DetailsPump As String, Account_CodeComm As String, _
    Account_Code As String, IsOther As Boolean
                    
                    
                    
                      Dim CIBAN As String, RecTime As Date, InvoiceTypeCodeID As Long, InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String, ActualDeliveryDate As Date
                      Dim LatestDeliveryDate As Date, paymentnote As String, PaymentMeansCode As String
   
                    
                    CIBAN = Trim(Rs3!CIBAN & "")
                    RecTime = IIf(IsNull(Rs3("RecTime").Value), Now, Rs3("RecTime").Value)
                    ActualDeliveryDate = IIf(IsNull(Rs3("ActualDeliveryDate").Value), Now, Rs3("ActualDeliveryDate").Value)
                    LatestDeliveryDate = IIf(IsNull(Rs3("LatestDeliveryDate").Value), Now, Rs3("LatestDeliveryDate").Value)
                    InvoiceTypeCodeID = Val(Rs3!InvoiceTypeCodeID & "")
                    InvoiceTypeCodename = Trim(Rs3!InvoiceTypeCodename & "")
                    DocumentCurrencyCode = Trim(Rs3!DocumentCurrencyCode & "")
                    TaxCurrencyCode = Trim(Rs3!TaxCurrencyCode & "")
                    paymentnote = Trim(Rs3!paymentnote & "")
                    PaymentMeansCode = Trim(Rs3!PaymentMeansCode & "")
                    
                    FromUserID = IIf(IsNull(Rs3("UserID").Value), 0, Rs3("UserID").Value)
                    FromEmp_ID = IIf(IsNull(Rs3("Emp_ID").Value), 0, Rs3("Emp_ID").Value)
                    FromStoreID = IIf(IsNull(Rs3("storeID").Value), 0, Rs3("storeID").Value)
                    CarTypeID = Val(Rs3!CarTypeID & "")
                    PlateNo = Trim(Rs3!PlateNo & "")
                    OilsTypesID = Val(Rs3!OilsTypesID & "")
                    YearFact = Val(Rs3!YearFact & "")
                    Shaseh = Trim(Rs3!Shaseh & "")
                    CarMeter = Trim(Rs3!CarMeter & "")
                    FixesAssetsID = Val(Rs3!FixesAssetsID & "")
                    ColorID2 = Val(Rs3!ColorID2 & "")
                    Chasee = Trim(Rs3!Chasee & "")
                    KM = Val(Rs3!KM & "")
                    PPointID = Val(Rs3!KM & "")
                    Phone2 = Trim(Rs3!Phone2 & "")
                    SupplerID = Val(Rs3!SupplerID & "")
                    Ser = Val(Rs3!Ser & "")
                    CarCurrentValue = Val(Rs3!CarCurrentValue & "")
                    CarPrevValue = Val(Rs3!CarPrevValue & "")
                    CarEnginoil = Val(Rs3!CarEnginoil & "")
                    CarGearOil = Val(Rs3!CarGearOil & "")
                    CarOilChangeDate = IIf(IsNull(Rs3("CarOilChangeDate").Value), Date, Rs3("CarOilChangeDate").Value)
                    
                    FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                    If FromTransaction_Type = 42 Then
                        s = "Select Code,CusName from TblCustemers where cusId =  " & FromCusID
                        Set rsCus = New ADODB.Recordset
                        rsCus.Open s, POSConnection, adOpenStatic, adLockReadOnly
                        If Not rsCus.EOF Then
                            s = "Select cusId,Code,CusName from TblCustemers where Code =  N'" & Trim(rsCus!Code & "") & "' and CusName =  N'" & Trim(rsCus!CusName & "") & "'"
                            Set rsCus2 = New ADODB.Recordset
                            rsCus2.Open s, Cn, adOpenStatic, adLockReadOnly
                            If Not rsCus2.EOF Then
                                FromCusID = Val(rsCus2!cusID & "")
                            End If
                            
                        End If
                        
                        
                    Else
                        FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                    End If
                    
                    FromBoxid = IIf(IsNull(Rs3("Boxid").Value), 0, Rs3("Boxid").Value)
                    POSBillType = IIf(IsNull(Rs3("POSBillType").Value), 0, Rs3("POSBillType").Value)
                    FromUserID = Val(Rs3!userID & "")
                    mTransaction_NetValue = Val(Rs3!Transaction_NetValue & "")
                    FromPaymentType = IIf(IsNull(Rs3("PaymentType").Value), 0, Rs3("PaymentType").Value)
                    FromBillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                    FromVATYou = IIf(IsNull(Rs3("VATYou").Value), 0, Rs3("VATYou").Value)
                    FromVAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
                    TypeInvoice = IIf(IsNull(Rs3("TypeInvoice").Value), 0, Rs3("TypeInvoice").Value)
                    '
                    BillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                    DepandToConv = True
                    Trans_DiscountType = IIf(IsNull(Rs3("Trans_DiscountType").Value), 0, Rs3("Trans_DiscountType").Value)
                    Trans_Discount = IIf(IsNull(Rs3("Trans_Discount").Value), 0, Rs3("Trans_Discount").Value)
                    TaxValue = IIf(IsNull(Rs3("TaxValue").Value), 0, Rs3("TaxValue").Value)
                    SaleType = IIf(IsNull(Rs3("SaleType").Value), 0, Rs3("SaleType").Value)
                    TaxAddValue = IIf(IsNull(Rs3("TaxAddValue").Value), 0, Rs3("TaxAddValue").Value)
                    
                    CashCustomerName = IIf(IsNull(Rs3("CashCustomerName").Value), "", Rs3("CashCustomerName").Value)
                    CashCustomerPhone = IIf(IsNull(Rs3("CashCustomerPhone").Value), "", Rs3("CashCustomerPhone").Value)
                    order_no = IIf(IsNull(Rs3("order_no").Value), "", Rs3("order_no").Value)
     
                    
                    FromBranchID = IIf(IsNull(Rs3("BranchID").Value), 0, Rs3("BranchID").Value)
                    fromTransaction_Serial = IIf(IsNull(Rs3("Transaction_Serial").Value), 0, Rs3("Transaction_Serial").Value)
                    
                    
                    FromNoteSerial1 = IIf(IsNull(Rs3("Noteserial1").Value), 0, Rs3("Noteserial1").Value)
                    FromNoteSerial = IIf(IsNull(Rs3("Noteserial").Value), 0, Rs3("Noteserial").Value)
                    'FromNoteId = IIf(IsNull(Rs3("NoteId").Value), 0, Rs3("NoteId").Value) ' —ﬁ„ ﬁÌœ «·›« Ê—…
                    
                    FromNots = IIf(IsNull(Rs3("Nots").Value), 0, Rs3("Nots").Value) '—ﬁ„ ”‰œ «·’—›
                    If FromTransaction_Type <> 42 Then
                        GetIssueData CDbl(Val(FromNots)), issueNoteid, issuenoteserial, issuenoteserial1
                    End If
                    FromNots2 = IIf(IsNull(Rs3("Nots2").Value), 0, Rs3("Nots2").Value)
                    FromTransaction_Date = IIf(IsNull(Rs3("Transaction_Date").Value), 0, Rs3("Transaction_Date").Value)
                    last_changed = IIf(IsNull(Rs3("last_changed").Value), Date, Rs3("last_changed").Value)
                    NetValue = IIf(IsNull(Rs3("NetValue").Value), 0, Rs3("NetValue").Value)
                    Transaction_Date = FromTransaction_Date
                    Transaction_Type = FromTransaction_Type
                    
                    Select Case Transaction_Type
                    Case 21
                        mNoteType = 170
                        mSanadNo = 7
                        CountSales = CountSales + 1
                    Case 19
                        mNoteType = 180
                        mSanadNo = 10
                        'CountSales =CountSales +1
                    Case 20
                        mNoteType = 160
                        mSanadNo = 9
                    
                    Case 5
                        mNoteType = 230
                        mSanadNo = 15
                        CountPurchaseReturn = CountPurchaseReturn + 1
                    Case 22
                        mNoteType = 150
                        mSanadNo = 6
                        CountPurchase = CountPurchase + 1
                    Case 9
                        mNoteType = 220
                        mSanadNo = 14
                        CountSalesReturn = CountSalesReturn + 1
                    Case 42
                        mNoteType = 0
                        mSanadNo = 42
                        CountSalesOfeers = CountSalesOfeers + 1
                    
                    End Select
                    isFoundData = True
                    
                    BranchID = FromBranchID
                    Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
                    'Transaction_Serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=" & Transaction_Type & ""))
                    'NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , Transaction_Type, , , , , , FromUserID)
                    NoteSerial1 = Rs3!NoteSerial1 & ""
                    If POSBillType = 0 Then
                        NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
                        NoteId = CStr(new_id("Notes", "NoteID", "", True))
                    End If
                    TransactionComment = " ›« Ê—… „‰ﬁÊ·… „‰ ﬁ«⁄œ…  " & POSname.Text & "   "
                    TransactionComment = TransactionComment & "  —ﬁ„ «·›« Ê—…  «·«’·Ì…" & FromNoteSerial1
                 '" & ServerDb & "
                 '   MsgBox TransactionComment
     
                    '*****************************************************
                    '*****************************************************
                    If Trim(NoteSerial) = "" Then NoteSerial = "0"
                    If Val(NoteId) = 0 Then NoteId = 0
                    'ÂÌœ— «·›« Ê—…
                    '*****************************************************************************************
                   

                    
                    
                    sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transactions]  (    "
                    sql = sql & "  Transaction_ID,Transaction_Date,TypeInvoice,"
                    sql = sql & "   Transaction_Serial ,"
                    sql = sql & "   Transaction_Type, "
                    sql = sql & "  PaymentType,"
                    sql = sql & "   CusID, StoreID, "
                    sql = sql & "  UserID, Emp_ID, "
                    sql = sql & "  BranchId, BoxID , "
                    sql = sql & "  BillBasedOn, VAT, "
                    sql = sql & "  VATYou, NoteSerial,"
                    sql = sql & "  NoteSerial1,NoteId,"
                    sql = sql & "  Copied,TransactionComment,"
                    sql = sql & " SessionCode,POSBillType, "
                    sql = sql & "  OldNoteserial1,OldNoteserial,"
                    sql = sql & " OldNoteId,OldTransaction_ID,"
                    
                    sql = sql & " Trans_DiscountType  ,"
                    sql = sql & " Trans_Discount   ,"
                    sql = sql & "TaxValue  ,"
                    sql = sql & " order_no  ,"
                    sql = sql & " SaleType  ,"
                    sql = sql & " CashCustomerName  ,"
                    sql = sql & "TaxAddValue  ,"
                    sql = sql & "CashCustomerPhone,last_changed ,NetValue,Transaction_NetValue,DepandToConv ,"
                    
                    
                    
                    sql = sql & "CarTypeID,"
                    sql = sql & "PlateNo,"
                    sql = sql & "OilsTypesID,"
                    sql = sql & "YearFact,"
                    sql = sql & "Shaseh,"
                    sql = sql & "CarMeter,"
                    sql = sql & "FixesAssetsID,"
                    sql = sql & "ColorID2,"
                    sql = sql & "KM,"
                    sql = sql & "Chasee,"
                    sql = sql & "PPointID,"
                    sql = sql & "Phone2,"
                    sql = sql & "SupplerID,"
                    sql = sql & "Ser,"
                    sql = sql & "CarCurrentValue,"
                    sql = sql & "CarPrevValue,"
                    sql = sql & "CarEnginoil,"
                    sql = sql & "CarGearOil,"
                    sql = sql & "CarOilChangeDate             ,   "
                    
                    sql = sql & "CIBAN             ,   "
                    sql = sql & "RecTime             ,   "
                    sql = sql & "ActualDeliveryDate             ,   "
                    sql = sql & "LatestDeliveryDate             ,   "
                    sql = sql & "InvoiceTypeCodeID             ,   "
                    sql = sql & "InvoiceTypeCodename             ,   "
                    sql = sql & "DocumentCurrencyCode             ,   "
                    sql = sql & "TaxCurrencyCode             ,   "
                    sql = sql & "paymentnote             ,   "
                    sql = sql & "PaymentMeansCode                "
                    sql = sql & ")"
                    
                    
              
                    
                    
                    sql = sql & "   values (" & Transaction_ID & "," & SQLDate(Transaction_Date, True) & "," & TypeInvoice & ","
                    sql = sql & FromNoteSerial1 & ","
                    sql = sql & Transaction_Type & ","
                    sql = sql & FromPaymentType & ","
                    sql = sql & FromCusID & "," & FromStoreID & ","
                    sql = sql & FromUserID & "," & FromEmp_ID & ","
                    sql = sql & FromBranchID & "," & FromBoxid
                    sql = sql & "," & BillBasedOn & "," & FromVAT & ","
                    sql = sql & FromVATYou & ","
                    sql = sql & NoteSerial & ",'" & FromNoteSerial1 & "'," & NoteId & ",1,'"
                    sql = sql & TransactionComment & "','" & SessionCode & "', " & POSBillType & " , '" & FromNoteSerial1 & "' , "
                    sql = sql & "'" & FromNoteSerial & "' , " & FromNoteId & " , " & FromTransaction_ID & " ,"
                    sql = sql & Trans_DiscountType & " ,"
                    sql = sql & Trans_Discount & " ,"
                    sql = sql & TaxValue & " ,"
                    sql = sql & "'" & order_no & "' ,"
                    sql = sql & SaleType & " ,"
                    sql = sql & "'" & CashCustomerName & "' ,"
                    sql = sql & TaxAddValue & " ,"
                    sql = sql & "'" & CashCustomerPhone & "' ," & SQLDate(last_changed, True) & "," & NetValue & "," & mTransaction_NetValue & "," & IIf(DepandToConv, 1, 0) & ", "
                    
                    
                    sql = sql & CarTypeID & ","
                    sql = sql & "'" & PlateNo & "',"
                    sql = sql & OilsTypesID & ","
                    sql = sql & YearFact & ","
                    sql = sql & "'" & Shaseh & "',"
                    sql = sql & "'" & CarMeter & "',"
                    sql = sql & FixesAssetsID & ","
                    sql = sql & ColorID2 & ","
                    sql = sql & KM & ","
                    sql = sql & "'" & Chasee & "',"
                    sql = sql & PPointID & ","
                    sql = sql & "'" & Phone2 & "',"
                    sql = sql & SupplerID & ","
                    sql = sql & Ser & ","
                    sql = sql & CarCurrentValue & ","
                    sql = sql & CarPrevValue & ","
                    sql = sql & CarEnginoil & ","
                    sql = sql & CarGearOil & ","
                    sql = sql & SQLDate(CarOilChangeDate, True) & ",                      "
                    sql = sql & "'" & CIBAN & "',"
                    sql = sql & SQLDate(RecTime, True) & ",                      "
                    sql = sql & SQLDate(ActualDeliveryDate, True) & ",                      "
                    sql = sql & SQLDate(LatestDeliveryDate, True) & ",                      "
                    sql = sql & InvoiceTypeCodeID & ","
                    sql = sql & "'" & InvoiceTypeCodename & "',"
                    sql = sql & "'" & DocumentCurrencyCode & "',"
                    sql = sql & "'" & TaxCurrencyCode & "',"
                    sql = sql & "'" & paymentnote & "',"
                    sql = sql & "'" & PaymentMeansCode & "'"
                    sql = sql & ")"
                    
                    
       



                    
                    
                    '   fromTransaction_Serial
                    Text1.Text = sql
                   ' Exit Sub
                   Text4 = ""
                       
     mLastStep = "Insert missing items into remote server"
mLastSQL = sql
                    Cn.Execute sql
                    Text4 = sql
                    
                    ' ›«’Ì· «·›« Ê—…
                    
                    
                    sql = " select * from Transaction_Details   where  Transaction_ID=" & FromTransaction_ID
                    Set rsDouble_Entry = New ADODB.Recordset
                    '
                    rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    Dim j As Double
                    For j = 1 To rsDouble_Entry.RecordCount
                        Item_ID = IIf(IsNull(rsDouble_Entry("Item_ID").Value), 0, rsDouble_Entry("Item_ID").Value)
                        ItemCase = IIf(IsNull(rsDouble_Entry("ItemCase").Value), 0, rsDouble_Entry("ItemCase").Value)
                        Quantity = IIf(IsNull(rsDouble_Entry("Quantity").Value), 0, rsDouble_Entry("Quantity").Value)
                        Price = IIf(IsNull(rsDouble_Entry("Price").Value), 0, rsDouble_Entry("Price").Value)
                        ItemDiscountType = IIf(IsNull(rsDouble_Entry("ItemDiscountType").Value), 0, rsDouble_Entry("ItemDiscountType").Value)
                        ItemDiscount = IIf(IsNull(rsDouble_Entry("ItemDiscount").Value), 0, rsDouble_Entry("ItemDiscount").Value)
                        ShowQty = IIf(IsNull(rsDouble_Entry("ShowQty").Value), 0, rsDouble_Entry("ShowQty").Value)
                        showPrice = IIf(IsNull(rsDouble_Entry("showPrice").Value), 0, rsDouble_Entry("showPrice").Value)
                        UnitID = IIf(IsNull(rsDouble_Entry("UnitId").Value), 0, rsDouble_Entry("UnitId").Value)
                        ColorID = IIf(IsNull(rsDouble_Entry("ColorID").Value), 0, rsDouble_Entry("ColorID").Value)
                        ItemSize = IIf(IsNull(rsDouble_Entry("ItemSize").Value), 0, rsDouble_Entry("ItemSize").Value)
                        ClassId = IIf(IsNull(rsDouble_Entry("ClassId").Value), 0, rsDouble_Entry("ClassId").Value)
                        mmVatyo = IIf(IsNull(rsDouble_Entry("Vatyo").Value), 0, rsDouble_Entry("Vatyo").Value)
                    
                        PumpId = Val(rsDouble_Entry!PumpId & "")
                        PrevQty = Val(rsDouble_Entry!PrevQty & "")
                        PrintName = Trim(rsDouble_Entry!PrintName & "")
                        Cash = Val(rsDouble_Entry!Cash & "")
                        Mada = Val(rsDouble_Entry!Mada & "")
                        Visa = Val(rsDouble_Entry!Visa & "")
                        Deferred = Val(rsDouble_Entry!Deferred & "")
                        AmountH = Val(rsDouble_Entry!AmountH & "")
                        AmountHComm = Val(rsDouble_Entry!AmountHComm & "")
                        DetailsPump = Trim(rsDouble_Entry!DetailsPump & "")
                        Account_CodeComm = Trim(rsDouble_Entry!Account_CodeComm & "")
                        Account_Code = Trim(rsDouble_Entry!Account_Code & "")
                        IsOther = Val(rsDouble_Entry!IsOther & "")
                        
                        sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transaction_Details]  (    "
                        sql = sql & "  Transaction_ID,  Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice,UnitId , ColorID, ItemSize, ClassId,SessionCode,Vatyo,"
                        
                        sql = sql & "  PumpId,"
                        sql = sql & "  PrevQty,"
                        sql = sql & "  PrintName,"
                        sql = sql & "  Cash,"
                        sql = sql & "  Mada,"
                        sql = sql & "  Visa,"
                        sql = sql & "  Deferred,"
                        sql = sql & "  AmountH,"
                        sql = sql & "  AmountHComm,"
                        sql = sql & "  DetailsPump,"
                        sql = sql & "  Account_CodeComm,"
                        sql = sql & "  Account_Code,"
                        sql = sql & "  IsOther"
                        
                        sql = sql & "  )"
                        sql = sql & "   values (" & Transaction_ID & "," & Item_ID & ", " & ItemCase & "," & Quantity & "," & Price & "," & ItemDiscountType & "," & ItemDiscount & "," & ShowQty & "," & showPrice
                        sql = sql & "," & UnitID & "," & ColorID & "," & ItemSize & "," & ClassId & "" & ",'" & SessionCode & "'," & mmVatyo & ","
                        
                        sql = sql & PumpId & ","
                        sql = sql & PrevQty & ","
                        sql = sql & "'" & PrintName & "',"
                        sql = sql & Cash & ","
                        sql = sql & Mada & ","
                        sql = sql & Visa & ","
                        sql = sql & Deferred & ","
                        sql = sql & AmountH & ","
                        sql = sql & AmountHComm & ","
                        sql = sql & "'" & DetailsPump & "',"
                        sql = sql & "'" & Account_CodeComm & "',"
                        sql = sql & "'" & Account_Code & "',"
                        sql = sql & IIf(IsOther, 1, 0) & ")"
                        
                        mLastStep = "Insert missing items into remote server"
mLastSQL = sql
                        Cn.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j
              '      MsgBox "3"
    '*********************** ›«’Ì· «·›«  ********************************************************
                    sql = " select * from TransactionValueAdded   where  Transaction_ID=" & FromTransaction_ID
                    Set rsDouble_Entry = New ADODB.Recordset
                    '
                    rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    'Dim j As Double
                    For j = 1 To rsDouble_Entry.RecordCount
                        ItemID = IIf(IsNull(rsDouble_Entry("ItemID").Value), 0, rsDouble_Entry("ItemID").Value)
                        Vatyo = IIf(IsNull(rsDouble_Entry("Vatyo").Value), 0, rsDouble_Entry("Vatyo").Value)
                        VAT = IIf(IsNull(rsDouble_Entry("Vat").Value), 0, rsDouble_Entry("Vat").Value)
                        Valu = IIf(IsNull(rsDouble_Entry("Valu").Value), 0, rsDouble_Entry("Valu").Value)
                        selectd = IIf(IsNull(rsDouble_Entry("selectd").Value), 0, rsDouble_Entry("selectd").Value)
                        
                        
                        sql = " INSERT INTO  [" & ServerDb & "].[dbo].[TransactionValueAdded]  (    "
                        sql = sql & "  Transaction_ID,  ItemID, Vatyo, VAT, Valu, selectd,Transaction_Type,SessionCode)"
                        sql = sql & "   values (" & Transaction_ID & "," & ItemID & ", " & Vatyo & "," & VAT & "," & Valu & "," & selectd & "," & Transaction_Type & " ,'" & SessionCode & "' )"
                        
                                                mLastStep = "Insert missing items into remote server"
mLastSQL = sql
                        Cn.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j
  
    '*********************** ›«’Ì· «·›«  ********************************************************
                    
     
    '*********************** ›«’Ì· «·‘»ﬂ… ********************************************************
                    If Transaction_Type = 21 Or Transaction_Type = 9 Then
                        sql = " select * from TblTransactionPayments   where  Transaction_ID=" & FromTransaction_ID
                        Set rsDouble_Entry = New ADODB.Recordset
                        '
                        Dim Recorddate As Date
                        rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                        'Dim j As Double
                        For j = 1 To rsDouble_Entry.RecordCount
                            BoxID = IIf(IsNull(rsDouble_Entry("boxid").Value), 0, rsDouble_Entry("boxid").Value)
                            Recorddate = IIf(IsNull(rsDouble_Entry("Recorddate").Value), 0, rsDouble_Entry("Recorddate").Value)
                            PointID = IIf(IsNull(rsDouble_Entry("PointID").Value), 0, rsDouble_Entry("PointID").Value)
                            CurrentCashireID = IIf(IsNull(rsDouble_Entry("CurrentCashireID").Value), 0, rsDouble_Entry("CurrentCashireID").Value)
                            PaymentID = IIf(IsNull(rsDouble_Entry("PaymentID").Value), 0, rsDouble_Entry("PaymentID").Value)
                            Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                            CardNo = IIf(IsNull(rsDouble_Entry("CardNo").Value), 0, rsDouble_Entry("CardNo").Value)
                            Effect = IIf(IsNull(rsDouble_Entry("Effect").Value), 0, rsDouble_Entry("Effect").Value)
                            
                            
                            sql = " INSERT INTO  [" & ServerDb & "].[dbo].[TblTransactionPayments]  (    "
                            sql = sql & "  Transaction_ID,  boxid, Recorddate, PointID, CurrentCashireID, PaymentID,Value,CardNo,Effect,SessionCode)"
                            sql = sql & "   values (" & Transaction_ID & "," & BoxID & ", " & SQLDate(Recorddate, True) & "," & PointID & "," & CurrentCashireID & "," & PaymentID & "," & Value & ",'" & CardNo & "'," & Effect & ",'" & SessionCode & "')"
                                                    mLastStep = "Insert missing items into remote server"
mLastSQL = sql
                            
                            Cn.Execute sql
                            rsDouble_Entry.MoveNext
                        Next j
                 '       MsgBox "5"
                    End If
'                    MsgBox "3"
    '*********************** ›«’Ì·  «·‘»ﬂ… ********************************************************
      
      
             
                'ﬁÌœ «·›« Ê—…
                 
                If POSBillType = 0 And Transaction_Type <> 42 Then
                 
                 
                    sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,Transaction_ID,UserID,SessionCode)"
                    sql = sql & " values( " & NoteId & ", " & SQLDate(Transaction_Date, True) & " ,  " & mNoteType & ", '" & NoteSerial & "', '" & NoteSerial1 & "'," & BranchID & "," & Transaction_ID & ",1,'" & SessionCode & "')"
                                            mLastStep = "Insert missing items into remote server"
mLastSQL = sql
                    Cn.Execute sql
                    DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                    
                    
                    'Dim rsDouble_Entry As ADODB.Recordset
                    Set rsDouble_Entry = New ADODB.Recordset
                    sql = " select * from DOUBLE_ENTREY_VOUCHERS   where   Notes_ID=" & FromNoteId
                    rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    Dim w As Double
                    For w = 1 To rsDouble_Entry.RecordCount
                        Account_Code = IIf(IsNull(rsDouble_Entry("Account_Code").Value), 0, rsDouble_Entry("Account_Code").Value)
                        Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                        Credit_Or_Debit = IIf(IsNull(rsDouble_Entry("Credit_Or_Debit").Value), 0, rsDouble_Entry("Credit_Or_Debit").Value)
                        Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                        Double_Entry_Vouchers_Description = IIf(IsNull(rsDouble_Entry("Double_Entry_Vouchers_Description").Value), 0, rsDouble_Entry("Double_Entry_Vouchers_Description").Value) & Chr(13) & "  ”‰œ ’—› " & TransactionComment
                        'RecordDate = IIf(IsNull(rsDouble_Entry("RecordDate").Value), 0, rsDouble_Entry("RecordDate").Value)
                        DEV_ID_Line_No = IIf(IsNull(rsDouble_Entry("DEV_ID_Line_No").Value), 0, rsDouble_Entry("DEV_ID_Line_No").Value)
                        branch_id = IIf(IsNull(rsDouble_Entry("branch_id").Value), 0, rsDouble_Entry("branch_id").Value)
                        sql = "  INSERT INTO [" & ServerDb & "].[dbo].[DOUBLE_ENTREY_VOUCHERS]([Double_Entry_Vouchers_ID], [DEV_ID_Line_No], [Account_Code], [Value], [Credit_Or_Debit], [Double_Entry_Vouchers_Description], [RecordDate], [Notes_ID] ,branch_id,UserID,Transaction_ID,SessionCode) "
                        sql = sql & " values (  " & DEVID & ", " & DEV_ID_Line_No & ", '" & Account_Code & "', " & Value & ", " & Credit_Or_Debit & ", '" & Double_Entry_Vouchers_Description & "',  " & SQLDate(Transaction_Date, True) & ", " & NoteId & " ," & branch_id & "," & 1 & "," & Transaction_ID & ",'" & SessionCode & "')"
                        Cn.Execute sql
                        
                        
                        rsDouble_Entry.MoveNext
                    Next w
               '     MsgBox "6"
                    '*****************************************************************
                '**********************************************************
                 End If
    
      
              '     GetIssueData CDbl(FromNots), issueNoteid, issuenoteserial, issuenoteserial1
                Dim mTransType2 As Integer
                If Transaction_Type = 21 Or Transaction_Type = 5 Then
                    mTransType2 = 19
                Else
                    mTransType2 = 20
                End If
                If chkDontCopyIss.Value = vbUnchecked And Transaction_Type <> 42 Then
                    CopyIssueTtransaction Transaction_ID, CStr(NoteSerial1), CDbl(Val(FromNots)), CDbl(mTransType2), issuenoteserial, issuenoteserial1, SessionCode
                End If
             
                Rs3.MoveNext
             
                lblCount.Caption = Val(lblCount.Caption) + 1
            Next i
   '      MsgBox "7"
         
        End If
     
        Rs3.Close
      'Sql = Sql & "[" & POSDb & "].dbo.Transactions"
      '„‰⁄ «·‰ﬁ· „—… «Œ—Ì
      
    
            sql = "update   [" & POSDb & "].dbo.Transactions" & "  set  Copied =1,SessionCode = '" & SessionCode & "' "
      sql = sql & "  Where Copied Is Null And "
      sql = sql & GetQuery
      '& "  and dbo.Transactions.Transaction_Date ='" & SQLDate(dbRecordDate.Value, False) & "'"
      
     POSConnection.Execute sql
    ' MsgBox "8"
    
'      sql = "update   [" & POSDb & "].dbo.Transaction_Details" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE  Copied is null    "
'     POSConnection.Execute sql
     
End If
  

If chkRec.Value = vbChecked Then
 sql = " select * From Notes where NoteType=4   and   Copied is null " ' ”‰œ«  ﬁ»÷
 
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    
     If Rs3.RecordCount > 0 Then
      
            For i = 1 To Rs3.RecordCount
                CashingType = IIf(IsNull(Rs3("CashingType").Value), 0, Rs3("CashingType").Value)
                          
               FromNoteId = IIf(IsNull(Rs3("NoteID").Value), 0, Rs3("NoteID").Value)
               FromBranchID = IIf(IsNull(Rs3("Branch_no").Value), 0, Rs3("Branch_no").Value)
               
              
                FromNoteSerial1 = IIf(IsNull(Rs3("Noteserial1").Value), 0, Rs3("Noteserial1").Value)
                FromNoteSerial = IIf(IsNull(Rs3("Noteserial").Value), 0, Rs3("Noteserial").Value)
                BranchID = FromBranchID
                NoteDate = IIf(IsNull(Rs3("NoteDate").Value), 0, Rs3("NoteDate").Value)
                'notedate = FromTransaction_Date
                NoteSerial = Notes_coding(BranchID, CDate(NoteDate))
                DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                'Dim notedate1 As Date
              
                NoteSerial1 = Voucher_coding(CInt(BranchID), CDate(NoteDate), 2, 4, , , , , , , , FromUserID)
         
                NoteId = CStr(new_id("Notes", "NoteID", "", True))
                TransactionComment = " ”‰œ ﬁ»÷  „‰ﬁÊ· „‰ ﬁ«⁄œ…  " & POSDb & "   "
                TransactionComment = TransactionComment & "  —ﬁ„ «·”‰œ  «·«’·Ì" & FromNoteSerial1
                TransactionComment = TransactionComment & "  —ﬁ„ «·ﬁÌœ  «·«’·Ì" & FromNoteSerial
   
                EmpId = IIf(IsNull(Rs3("EmpId").Value), 0, Rs3("EmpId").Value)
                VAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
                person = IIf(IsNull(Rs3("person").Value), 0, Rs3("person").Value)
                NCashingType = IIf(IsNull(Rs3("NCashingType").Value), 0, Rs3("NCashingType").Value)
                Status = IIf(IsNull(Rs3("Status").Value), 0, Rs3("Status").Value)
                Note_Value = IIf(IsNull(Rs3("Note_Value").Value), 0, Rs3("Note_Value").Value)
                BankName = IIf(IsNull(Rs3("BankName").Value), 0, Rs3("BankName").Value)
                Remark = IIf(IsNull(Rs3("Remark").Value), 0, Rs3("Remark").Value)
                cusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                NoteCashingType = IIf(IsNull(Rs3("NoteCashingType").Value), 0, Rs3("NoteCashingType").Value)
                BoxID = IIf(IsNull(Rs3("BoxID").Value), "Null", Rs3("BoxID").Value)
                ChqueNum = IIf(IsNull(Rs3("ChqueNum").Value), 0, Rs3("ChqueNum").Value)
                DueDate = IIf(IsNull(Rs3("DueDate").Value), 0, Rs3("DueDate").Value)
                ChequeBoxID = IIf(IsNull(Rs3("ChequeBoxID").Value), 0, Rs3("ChequeBoxID").Value)
                BankID = IIf(IsNull(Rs3("BankID").Value), "Null", Rs3("BankID").Value)
                TotalNotesValue = IIf(IsNull(Rs3("TotalNotesValue").Value), 0, Rs3("TotalNotesValue").Value)
                
                sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,UserID,CashingType,EmpId,VAT"
                 sql = sql & ",NCashingType, Status,Note_Value,BankName,Remark,CusID,NoteCashingType,BoxID,ChqueNum,DueDate,ChequeBoxID,BankID,TotalNotesValue,copied,SessionCode )"
                 sql = sql & " values( " & NoteId & ", " & SQLDate(CDate(NoteDate), True) & " , 4, " & NoteSerial & ", " & NoteSerial1 & "," & BranchID & ",1," & CashingType & "," & EmpId & "," & VAT
                 sql = sql & "," & NCashingType & ", " & Status & "," & Note_Value & ",'" & BankName & "','" & Remark & "'," & cusID & "," & NoteCashingType & "," & BoxID & "," & ChqueNum & "," & SQLDate(CDate(Date), True) & "," & ChequeBoxID & "," & BankID & "," & TotalNotesValue & ",1,'" & SessionCode & "')"
                 Cn.Execute sql
                 DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                 
                 
                 'Dim rsDouble_Entry As ADODB.Recordset
                  Set rsDouble_Entry = New ADODB.Recordset
                     sql = " select * from DOUBLE_ENTREY_VOUCHERS   where   Notes_ID=" & FromNoteId
                   rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    'Dim w As Double
                    For w = 1 To rsDouble_Entry.RecordCount
                    Account_Code = IIf(IsNull(rsDouble_Entry("Account_Code").Value), 0, rsDouble_Entry("Account_Code").Value)
                    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                    Credit_Or_Debit = IIf(IsNull(rsDouble_Entry("Credit_Or_Debit").Value), 0, rsDouble_Entry("Credit_Or_Debit").Value)
                    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                    Double_Entry_Vouchers_Description = IIf(IsNull(rsDouble_Entry("Double_Entry_Vouchers_Description").Value), 0, rsDouble_Entry("Double_Entry_Vouchers_Description").Value) & Chr(13) & "    " & TransactionComment
                    Recorddate = IIf(IsNull(rsDouble_Entry("RecordDate").Value), 0, rsDouble_Entry("RecordDate").Value)
                    DEV_ID_Line_No = IIf(IsNull(rsDouble_Entry("DEV_ID_Line_No").Value), 0, rsDouble_Entry("DEV_ID_Line_No").Value)
                    branch_id = IIf(IsNull(rsDouble_Entry("branch_id").Value), 0, rsDouble_Entry("branch_id").Value)
                    sql = "  INSERT INTO [" & ServerDb & "].[dbo].[DOUBLE_ENTREY_VOUCHERS]([Double_Entry_Vouchers_ID], [DEV_ID_Line_No], [Account_Code], [Value], [Credit_Or_Debit], [Double_Entry_Vouchers_Description], [RecordDate], [Notes_ID] ,branch_id,UserID ,SessionCode ) "
                    sql = sql & " values (  " & DEVID & ", " & DEV_ID_Line_No & ", '" & Account_Code & "', " & Value & ", " & Credit_Or_Debit & ", '" & Double_Entry_Vouchers_Description & "',  " & SQLDate(Recorddate, True) & ", " & NoteId & " ," & branch_id & ", 1,'" & SessionCode & "')"
                    Cn.Execute sql
                
                
                    rsDouble_Entry.MoveNext
                    Next w
                
                
                '*********************** ›«’Ì· «·‘Ìﬂ… ··ﬁ»÷ ********************************************************
                    sql = " select * from TblMultuPayment   where  NoteID=" & FromNoteId
                    Set rsDouble_Entry = New ADODB.Recordset
                    '
                     
                    rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    'Dim j As Double
                    For j = 1 To rsDouble_Entry.RecordCount
                '     NoteId = IIf(IsNull(rsDouble_Entry("NoteId").Value), 0, rsDouble_Entry("NoteId").Value)
                        PaymentID = IIf(IsNull(rsDouble_Entry("PaymentID").Value), 0, rsDouble_Entry("PaymentID").Value)
                        Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                        CardNo = IIf(IsNull(rsDouble_Entry("CardNo").Value), " ", rsDouble_Entry("CardNo").Value)
                        maxvalue = IIf(IsNull(rsDouble_Entry("maxvalue").Value), 0, rsDouble_Entry("maxvalue").Value)
                        sql = " INSERT INTO  [" & ServerDb & "].[dbo].[TblMultuPayment]  (    "
                        sql = sql & "  NoteId,   PaymentID, Value, CardNo, maxvalue ,SessionCode )"
                        sql = sql & "   values (" & NoteId & ", " & PaymentID & "," & Value & ",'" & CardNo & "'," & maxvalue & ",'" & SessionCode & "')"
                        
                        
                        Cn.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j

'*********************** ··ﬁ»÷  ›«’Ì·  «·‘»ﬂ… ********************************************************
  
  
 
            Next i
            
      End If
        sql = "update   [" & POSDb & "].dbo.Notes" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE   NoteType=4   and   Copied is null  "
        sql = sql & " and dbo.Notes.NoteDate ='" & SQLDate(dbRecordDate.Value, False) & "'"
        POSConnection.Execute sql
     '   MsgBox "9"
        
'        sql = "update   [" & POSDb & "].dbo.DOUBLE_ENTREY_VOUCHERS" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE  Copied is null   "
'        POSConnection.Execute sql


 End If
If isFoundData Then
     Dim rsOffline As New ADODB.Recordset
    Dim mEndTime22 As String
    mEndTime22 = Now
    s = "Select * from TblOffline where 1 = -1"
    rsOffline.Open s, Cn, adOpenKeyset, adLockOptimistic
    'MsgBox s
    rsOffline.AddNew
    'MsgBox s & "Save"
    'rsOffline!Id = mMaxId
    rsOffline!Recorddate = Date
    rsOffline!EndTime = mEndTime22
    rsOffline!StartTime = mTimeStart
    rsOffline!SessionCode = SessionCode
    rsOffline!POSname = POSlServer
    
    rsOffline!CountSalesOfeers = CountSalesOfeers
    rsOffline!CountSales = CountSales
    rsOffline!CountSalesReturn = CountSalesReturn
    rsOffline!CountPurchase = CountPurchase
    rsOffline!CountPurchaseReturn = CountPurchaseReturn
    rsOffline!CountRec = CountRec
    rsOffline.Update
    
    Cn.CommitTrans
    BeginTrans = False
End If

'MsgBox " „ ‰ﬁ· «·»Ì«‰« "

 
    





'Dim mMaxId As Long
's = "Select Max(Id) as MaxID  from TblOffline"
'rsOffline.Open s, Cn, adOpenKeyset, adLockOptimistic
'mMaxId = 1
'If Not rsOffline.EOF Then
'    mMaxId = Val(rsOffline!MaxID & "") + 1
'
'End If
'rsOffline.Close

lblWait.Visible = False

txtEndTime = mEndTime22
txtCountSalesReturn = CountSalesReturn
txtCountSales = CountSales
txtCountSalesOfeers = CountSalesOfeers
Exit Sub



ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'Resume Next
'MsgBox "Done"
'////////////////////////////////////////copy Sales Transactions

End Sub

Private Sub Command10_Click()
Dim sql As String
Dim sqlPart1 As String, sqlPart2 As String, sqlPart3 As String
Dim rs As ADODB.Recordset
Dim rsDetails As ADODB.Recordset
Dim nextTransactionID As Long
Dim currentTransactionID As Long
Dim sqlValues1 As String
Dim sqlValues2 As String
Dim sqlValues3 As String

On Error Resume Next
  Dim mPosD As String
            Dim mServerD As String
            mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
            mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
            mServerD = "RemoteServer10"
           mServerD = "[RemoteServer10].Byte.dbo."
lblWait.Caption = "Ì „ «·«‰ ‰ﬁ· ”‰œ«  «· ÕÊÌ·«  «·„Œ“‰Ì…"
  Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If SysSQLServerType = 1 Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'            "Persist Security Info=False;Initial Catalog=" & POSDb & _
'            ";Data Source=" & POSlServer & ";Port=1433"
'
        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
        ElseIf SysSQLServerType = 2 Then
             If SysSQLServerTypeTechnical = "0" Then
             .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                "Persist Security Info=False;Initial Catalog=" & POSDb & _
                ";Data Source=" & POSlServer & ";Port=1433"
              Else
                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer 'SysSQLServerName
            End If
        End If
       '   Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Adnan;Data Source=WAELPC\SQLEXPRESS;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=WAELPC;Use Encryption for Data=False;Tag with column collation when possible=False;

        '.Open
    End With
    
     POSConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
      POSConnection.Open
    
' Ã·» √ﬂ»— Transaction_ID Õ«·Ì „‰ ÃœÊ· Transactions ›Ì «·‰ﬁÿ… · Ê·Ìœ «·—ﬁ„ «·ÃœÌœ
sql = "SELECT ISNULL(MAX(Transaction_ID), 0) AS MaxTransactionID FROM [" & POSDb & "].[dbo].[Transactions]"
Set rs = New ADODB.Recordset
rs.Open sql, POSConnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
    nextTransactionID = rs("MaxTransactionID").Value + 1
Else
    nextTransactionID = 1 ' »œ«Ì… „‰ 1 ≈–« ·„ Ìﬂ‰ Â‰«ﬂ »Ì«‰« 
End If
rs.Close


Dim BeginTrans  As Boolean
 POSConnection.BeginTrans
                BeginTrans = True
' Ã·» «·»Ì«‰«  „‰ ÃœÊ· Transactions ›Ì «·”Ì—›—
sql = "SELECT * FROM " & mServerD & "[Transactions] T2 " & _
      "WHERE Transaction_Type in (10,11) and NOT EXISTS (SELECT 1 FROM " & mPosD & "[Transactions] T1 WHERE T1.OldTransaction_ID = T2.Transaction_ID)"
Set rs = New ADODB.Recordset
'Cn.Open
rs.Open sql, POSConnection, adOpenStatic, adLockReadOnly

' „⁄«·Ã… ﬂ· ’› · Ê·Ìœ Transaction_ID ÃœÌœ
Do While Not rs.EOF
    '  ⁄ÌÌ‰ Transaction_ID «·ÃœÌœ
    currentTransactionID = nextTransactionID
    nextTransactionID = nextTransactionID + 1 '  ÕœÌÀ «·—ﬁ„ ··’› «· «·Ì

    '  ﬁ”Ì„ «” ⁄·«„ «·≈÷«›… ≈·Ï √Ã“«¡
'    sqlPart1 = "INSERT INTO " & mPosD & "[Transactions] (Transaction_ID, OldTransaction_ID, Transaction_Date, TypeInvoice, " & _
'               "Transaction_Serial, Transaction_Type, PaymentType, CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, "
'    sqlPart2 = "NoteSerial, NoteSerial1, NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, OldNoteserial, " & _
'               "OldNoteId, Trans_DiscountType, Trans_Discount, TaxValue, order_no, SaleType, CashCustomerName, TaxAddValue, CashCustomerPhone, "
'    sqlPart3 = "last_changed, NetValue, Transaction_NetValue, DepandToConv, CarTypeID, PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, " & _
'               "FixesAssetsID, ColorID2, KM, Chasee, PPointID, Phone2, SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, CarGearOil, CarOilChangeDate) VALUES "
'
'    ' »‰«¡ «” ⁄·«„ «·≈÷«›…
'    sql = sqlPart1 & sqlPart2 & sqlPart3
'    sql = sql & "(" & currentTransactionID & ", " & rs("Transaction_ID") & ", '" & rs("Transaction_Date") & "', " & _
'          rs("TypeInvoice") & ", '" & rs("Transaction_Serial") & "', " & rs("Transaction_Type") & ", " & rs("PaymentType") & ", " & _
'          rs("CusID") & ", " & rs("StoreID") & ", " & rs("UserID") & ", " & rs("Emp_ID") & ", " & rs("BranchId") & ", " & _
'          rs("BoxID") & ", " & rs("BillBasedOn") & ", " & rs("VAT") & ", " & rs("VATYou") & ", " & rs("NoteSerial") & ", " & _
'          "'" & rs("NoteSerial1") & "', " & rs("NoteId") & ", " & rs("Copied") & ", '" & rs("TransactionComment") & "', " & _
'          "'" & rs("SessionCode") & "', " & rs("POSBillType") & ", '" & rs("OldNoteserial1") & "', '" & rs("OldNoteserial") & "', " & _
'          rs("OldNoteId") & ", " & rs("Trans_DiscountType") & ", " & rs("Trans_Discount") & ", " & rs("TaxValue") & ", '" & _
'          rs("order_no") & "', " & rs("SaleType") & ", '" & rs("CashCustomerName") & "', " & rs("TaxAddValue") & ", '" & _
'          rs("CashCustomerPhone") & "', '" & rs("last_changed") & "', " & rs("NetValue") & ", " & rs("Transaction_NetValue") & ", " & _
'          rs("DepandToConv") & ", " & rs("CarTypeID") & ", '" & rs("PlateNo") & "', " & rs("OilsTypesID") & ", " & rs("YearFact") & ", " & _
'          "'" & rs("Shaseh") & "', '" & rs("CarMeter") & "', " & rs("FixesAssetsID") & ", " & rs("ColorID2") & ", " & rs("KM") & ", '" & _
'          rs("Chasee") & "', " & rs("PPointID") & ", '" & rs("Phone2") & "', " & rs("SupplerID") & ", " & rs("Ser") & ", " & _
'          rs("CarCurrentValue") & ", " & rs("CarPrevValue") & ", " & rs("CarEnginoil") & ", " & rs("CarGearOil") & ", '" & _
'          rs("CarOilChangeDate") & "')"

    '  ‰›Ì– «·«” ⁄·«„
    
    ' «·Ã“¡ «·√Ê· „‰ «·«” ⁄·«„ («·√⁄„œ…)
sqlPart1 = "INSERT INTO " & mPosD & "[Transactions] (Transaction_ID, OldTransaction_ID, Transaction_Date,  " & _
           "Transaction_Serial, Transaction_Type,  CusID, StoreID, UserID, Emp_ID, BranchId,  VAT, VATYou, "

sqlPart2 = "NoteSerial, NoteSerial1,  Copied,  SessionCode,  OldNoteserial1,  " & _
            "order_no, TaxAddValue,  "

sqlPart3 = "NetValue, Transaction_NetValue,  " & _
           " Ser) VALUES ("




sqlPart1 = "INSERT INTO " & mPosD & "[Transactions] (Transaction_ID, OldTransaction_ID, Transaction_Date, " & _
           "Transaction_Serial, Transaction_Type,  CusID, StoreID, UserID, Emp_ID, BranchId,  VAT, VATYou, "

sqlPart2 = "NoteSerial, NoteSerial1,  Copied,  SessionCode,  OldNoteserial1,  " & _
             "  order_no,  TaxAddValue,  "

sqlPart3 = " NetValue, Transaction_NetValue,  " & _
           " Ser) VALUES ("

' «·Ã“¡ «·√Ê· „‰ «·ﬁÌ„




sqlValues1 = currentTransactionID & ", "
'''
sqlValues1 = sqlValues1 & IIf(IsNull(rs("Transaction_ID")) Or rs("Transaction_ID") = "", "NULL", rs("Transaction_ID")) & ", "
'sqlValues1 = sqlValues1 & IIf(IsNull(rs("Transaction_Date")) Or rs("Transaction_Date") = "", "NULL", "'" & rs("Transaction_Date") & "'") & ", "
If IsDate(rs("Transaction_Date")) Then
    sqlValues1 = sqlValues1 & "'" & Format(rs("Transaction_Date"), "YYYY-MM-DD") & "', "
Else
    sqlValues1 = sqlValues1 & "NULL, "
End If

sqlValues1 = sqlValues1 & IIf(IsNull(rs("TypeInvoice")) Or rs("TypeInvoice") = "", "NULL", rs("TypeInvoice")) & ", "
sqlValues1 = sqlValues1 & IIf(IsNull(rs("Transaction_Serial")) Or rs("Transaction_Serial") = "", "NULL", "'" & rs("Transaction_Serial") & "'") & ", "
sqlValues1 = sqlValues1 & IIf(IsNull(rs("Transaction_Type")) Or rs("Transaction_Type") = "", "NULL", rs("Transaction_Type")) & ", "
sqlValues1 = sqlValues1 & IIf(IsNull(rs("PaymentType")) Or rs("PaymentType") = "", "NULL", rs("PaymentType")) & ", "
sqlValues1 = sqlValues1 & IIf(IsNull(rs("CusID")) Or rs("CusID") = "", "NULL", rs("CusID")) & ", "
sqlValues1 = sqlValues1 & IIf(IsNull(rs("StoreID")) Or rs("StoreID") = "", "NULL", rs("StoreID"))

'Debug.Print "sqlValues1 After Assembly: " & sqlValues1

' «·Ã“¡ «·À«‰Ì „‰ «·ﬁÌ„
sqlValues2 = IIf(IsNull(rs("UserID")) Or rs("UserID") = "", "NULL", rs("UserID")) & ", " & _
             IIf(IsNull(rs("Emp_ID")) Or rs("Emp_ID") = "", "NULL", rs("Emp_ID")) & ", " & _
             IIf(IsNull(rs("BranchId")) Or rs("BranchId") = "", "NULL", rs("BranchId")) & ", " & _
             IIf(IsNull(rs("BoxID")) Or rs("BoxID") = "", "NULL", rs("BoxID")) & ", " & _
             IIf(IsNull(rs("BillBasedOn")) Or rs("BillBasedOn") = "", "NULL", rs("BillBasedOn")) & ", " & _
             IIf(IsNull(rs("VAT")) Or rs("VAT") = "", "NULL", rs("VAT")) & ", " & _
             IIf(IsNull(rs("VATYou")) Or rs("VATYou") = "", "NULL", rs("VATYou")) & ", " & _
             IIf(IsNull(rs("NoteSerial")) Or rs("NoteSerial") = "", "NULL", rs("NoteSerial")) & ", " & _
             IIf(IsNull(rs("NoteSerial1")) Or rs("NoteSerial1") = "", "NULL", "'" & rs("NoteSerial1") & "'") & ", "

' «·Ã“¡ «·À«·À „‰ «·ﬁÌ„
sqlValues3 = IIf(IsNull(rs("NoteId")) Or rs("NoteId") = "", "NULL", rs("NoteId")) & ", " & _
             IIf(IsNull(rs("Copied")) Or rs("Copied") = "", "NULL", rs("Copied")) & ", " & _
             IIf(IsNull(rs("TransactionComment")) Or rs("TransactionComment") = "", "NULL", "'" & rs("TransactionComment") & "'") & ", " & _
             IIf(IsNull(rs("SessionCode")) Or rs("SessionCode") = "", "NULL", "'" & rs("SessionCode") & "'") & ", " & _
             IIf(IsNull(rs("POSBillType")) Or rs("POSBillType") = "", "NULL", rs("POSBillType")) & ", " & _
             IIf(IsNull(rs("OldNoteserial1")) Or rs("OldNoteserial1") = "", "NULL", "'" & rs("OldNoteserial1") & "'") & ", " & _
             IIf(IsNull(rs("OldNoteserial")) Or rs("OldNoteserial") = "", "NULL", "'" & rs("OldNoteserial") & "'") & ", " & _
             IIf(IsNull(rs("OldNoteId")) Or rs("OldNoteId") = "", "NULL", rs("OldNoteId")) & ", " & _
             IIf(IsNull(rs("Trans_DiscountType")) Or rs("Trans_DiscountType") = "", "NULL", rs("Trans_DiscountType")) & ", "

' «·Ã“¡ «·—«»⁄ „‰ «·ﬁÌ„
sqlValues4 = IIf(IsNull(rs("Trans_Discount")) Or rs("Trans_Discount") = "", "NULL", rs("Trans_Discount")) & ", " & _
             IIf(IsNull(rs("TaxValue")) Or rs("TaxValue") = "", "NULL", rs("TaxValue")) & ", " & _
             IIf(IsNull(rs("order_no")) Or rs("order_no") = "", "NULL", "'" & rs("order_no") & "'") & ", " & _
             IIf(IsNull(rs("SaleType")) Or rs("SaleType") = "", "NULL", rs("SaleType")) & ", " & _
             IIf(IsNull(rs("CashCustomerName")) Or rs("CashCustomerName") = "", "NULL", "'" & rs("CashCustomerName") & "'") & ", " & _
             IIf(IsNull(rs("TaxAddValue")) Or rs("TaxAddValue") = "", "NULL", rs("TaxAddValue")) & ", " & _
             IIf(IsNull(rs("CashCustomerPhone")) Or rs("CashCustomerPhone") = "", "NULL", "'" & rs("CashCustomerPhone") & "'") & ", " & _
             IIf(IsNull(rs("last_changed")) Or rs("last_changed") = "", "NULL", "'" & rs("last_changed") & "'") & ", " & _
             IIf(IsNull(rs("NetValue")) Or rs("NetValue") = "", "NULL", rs("NetValue")) & ", "

' «·Ã“¡ «·Œ«„” „‰ «·ﬁÌ„
sqlValues5 = IIf(IsNull(rs("Transaction_NetValue")) Or rs("Transaction_NetValue") = "", "NULL", rs("Transaction_NetValue")) & ", " & _
             IIf(IsNull(rs("DepandToConv")) Or rs("DepandToConv") = "", "NULL", rs("DepandToConv")) & ", " & _
             IIf(IsNull(rs("CarTypeID")) Or rs("CarTypeID") = "", "NULL", rs("CarTypeID")) & ", " & _
             IIf(IsNull(rs("PlateNo")) Or rs("PlateNo") = "", "NULL", "'" & rs("PlateNo") & "'") & ", " & _
             IIf(IsNull(rs("OilsTypesID")) Or rs("OilsTypesID") = "", "NULL", rs("OilsTypesID")) & ", " & _
             IIf(IsNull(rs("YearFact")) Or rs("YearFact") = "", "NULL", rs("YearFact")) & ", " & _
             IIf(IsNull(rs("Shaseh")) Or rs("Shaseh") = "", "NULL", "'" & rs("Shaseh") & "'") & ", " & _
             IIf(IsNull(rs("CarMeter")) Or rs("CarMeter") = "", "NULL", "'" & rs("CarMeter") & "'") & ", " & _
             IIf(IsNull(rs("FixesAssetsID")) Or rs("FixesAssetsID") = "", "NULL", rs("FixesAssetsID")) & ", "

' «·Ã“¡ «·√ŒÌ— „‰ «·ﬁÌ„
sqlValues6 = IIf(IsNull(rs("ColorID2")) Or rs("ColorID2") = "", "NULL", rs("ColorID2")) & ", " & _
             IIf(IsNull(rs("KM")) Or rs("KM") = "", "NULL", rs("KM")) & ", " & _
             IIf(IsNull(rs("Chasee")) Or rs("Chasee") = "", "NULL", "'" & rs("Chasee") & "'") & ", " & _
             IIf(IsNull(rs("PPointID")) Or rs("PPointID") = "", "NULL", rs("PPointID")) & ", " & _
             IIf(IsNull(rs("Phone2")) Or rs("Phone2") = "", "NULL", "'" & rs("Phone2") & "'") & ", " & _
             IIf(IsNull(rs("SupplerID")) Or rs("SupplerID") = "", "NULL", rs("SupplerID")) & ", " & _
             IIf(IsNull(rs("Ser")) Or rs("Ser") = "", "NULL", rs("Ser")) & ", " & _
             IIf(IsNull(rs("CarCurrentValue")) Or rs("CarCurrentValue") = "", "NULL", rs("CarCurrentValue")) & ", " & _
             IIf(IsNull(rs("CarPrevValue")) Or rs("CarPrevValue") = "", "NULL", rs("CarPrevValue")) & ", " & _
             IIf(IsNull(rs("CarEnginoil")) Or rs("CarEnginoil") = "", "NULL", rs("CarEnginoil")) & ", " & _
             IIf(IsNull(rs("CarGearOil")) Or rs("CarGearOil") = "", "NULL", rs("CarGearOil")) & ", " & _
             IIf(IsNull(rs("CarOilChangeDate")) Or rs("CarOilChangeDate") = "", "NULL", "'" & rs("CarOilChangeDate") & "'") & ")"

' œ„Ã «·«” ⁄·«„ «·ﬂ«„·
'sql = sqlPart1 & sqlPart2 & sqlPart3 & sqlValues1 & sqlValues2 & sqlValues3 & sqlValues4 & sqlValues5 & sqlValues6

    
    
sqlPart1 = "INSERT INTO [Transactions] (Transaction_ID, OldTransaction_ID, Transaction_Date, " & _
           "Transaction_Serial, Transaction_Type,  CusID, StoreID, UserID, Emp_ID, BranchId,  VAT, VATYou, "

sqlPart2 = "NoteSerial, NoteSerial1,  Copied,  SessionCode,  OldNoteserial1,  " & _
            "  order_no,  TaxAddValue,  "

sqlPart3 = " NetValue, Transaction_NetValue,  " & _
           " Ser) VALUES ("

' ≈⁄œ«œ «·ﬁÌ„ »ÕÌÀ  ÿ«»ﬁ «·√⁄„œ…
sqlValues1 = currentTransactionID & ", " & _
             IIf(IsNull(rs("Transaction_ID")) Or rs("Transaction_ID") = "", "NULL", rs("Transaction_ID")) & ", " & _
             IIf(IsNull(rs("Transaction_Date")) Or rs("Transaction_Date") = "", "NULL", "'" & Format(rs("Transaction_Date"), "YYYY-MM-DD") & "'") & ", " & _
             IIf(IsNull(rs("Transaction_Serial")) Or rs("Transaction_Serial") = "", "NULL", "'" & rs("Transaction_Serial") & "'") & ", " & _
             IIf(IsNull(rs("Transaction_Type")) Or rs("Transaction_Type") = "", "NULL", rs("Transaction_Type")) & ", " & _
             IIf(IsNull(rs("CusID")) Or rs("CusID") = "", "NULL", rs("CusID")) & ", " & _
             IIf(IsNull(rs("StoreID")) Or rs("StoreID") = "", "NULL", rs("StoreID")) & ", " & _
             IIf(IsNull(rs("UserID")) Or rs("UserID") = "", "NULL", rs("UserID")) & ", " & _
             IIf(IsNull(rs("Emp_ID")) Or rs("Emp_ID") = "", "NULL", rs("Emp_ID")) & ", "

sqlValues2 = IIf(IsNull(rs("BranchId")) Or rs("BranchId") = "", "NULL", rs("BranchId")) & ", " & _
             IIf(IsNull(rs("VAT")) Or rs("VAT") = "", "NULL", rs("VAT")) & ", " & _
             IIf(IsNull(rs("VATYou")) Or rs("VATYou") = "", "NULL", rs("VATYou")) & ", " & _
             IIf(IsNull(rs("NoteSerial")) Or rs("NoteSerial") = "", "NULL", rs("NoteSerial")) & ", " & _
             IIf(IsNull(rs("NoteSerial1")) Or rs("NoteSerial1") = "", "NULL", "'" & rs("NoteSerial1") & "'") & ", " & _
             IIf(IsNull(rs("Copied")) Or rs("Copied") = "", "NULL", rs("Copied")) & ", " & _
             IIf(IsNull(rs("SessionCode")) Or rs("SessionCode") = "", "NULL", "'" & rs("SessionCode") & "'") & ", "

sqlValues3 = IIf(IsNull(rs("OldNoteserial1")) Or rs("OldNoteserial1") = "", "NULL", "'" & rs("OldNoteserial1") & "'") & ", " & _
             IIf(IsNull(rs("order_no")) Or rs("order_no") = "", "NULL", "'" & rs("order_no") & "'") & ", " & _
             IIf(IsNull(rs("TaxAddValue")) Or rs("TaxAddValue") = "", "NULL", rs("TaxAddValue")) & ", " & _
             IIf(IsNull(rs("NetValue")) Or rs("NetValue") = "", "NULL", rs("NetValue")) & ", " & _
             IIf(IsNull(rs("Transaction_NetValue")) Or rs("Transaction_NetValue") = "", "NULL", rs("Transaction_NetValue")) & ", " & _
             IIf(IsNull(rs("Ser")) Or rs("Ser") = "", "NULL", rs("Ser")) & ")"

' œ„Ã «·‰’Ê’ «·‰Â«∆Ì…
sql = sqlPart1 & sqlPart2 & sqlPart3 & sqlValues1 & sqlValues2 & sqlValues3
                            mLastStep = "Insert missing items into remote server"
mLastSQL = sql
    POSConnection.Execute sql
    If Err.Number <> 0 Then
        frmPopup.ShowMessage "Error in Transactions: " & Err.Description
        Err.Clear
    End If

    ' Ã·» »Ì«‰«  Transaction_Details «·„— »ÿ…
    sql = "SELECT * FROM [" & ServerDb & "].[dbo].[Transaction_Details] TD " & _
          "WHERE TD.Transaction_ID = " & rs("Transaction_ID")
    Set rsDetails = New ADODB.Recordset
    rsDetails.Open sql, Cn, adOpenStatic, adLockReadOnly

    ' „⁄«·Ã… ﬂ· ’› „‰ Transaction_Details
    Do While Not rsDetails.EOF
        sql = "INSERT INTO  [Transaction_Details] (Transaction_ID, Item_ID, ItemCase, Quantity, Price, " & _
              "ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, ColorID, ItemSize, ClassId, SessionCode, Vatyo, " & _
              "PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, AmountHComm, DetailsPump, Account_CodeComm, " & _
              "Account_Code, IsOther) VALUES (" & _
              currentTransactionID & ", " & rsDetails("Item_ID") & ", " & Val(rsDetails("ItemCase") & "") & ", " & _
              Val(rsDetails("Quantity") & "") & ", " & Val(rsDetails("Price") & "") & ", " & Val(rsDetails("ItemDiscountType") & "") & ", " & _
              Val(rsDetails("ItemDiscount") & "") & ", " & Val(rsDetails("ShowQty") & "") & ", " & Val(rsDetails("showPrice") & "") & ", " & _
              Val(rsDetails("UnitId") & "") & ", " & Val(rsDetails("ColorID") & "") & ", " & Val(rsDetails("ItemSize") & "") & ", " & _
              Val(rsDetails("ClassId") & "") & ", '" & Trim(rsDetails("SessionCode") & "") & "', " & Val(rsDetails("Vatyo") & "") & ", " & _
              Val(rsDetails("PumpId") & "") & ", " & Val(rsDetails("PrevQty") & "") & ", '" & Trim(rsDetails("PrintName") & "") & "', " & _
              Val(rsDetails("Cash") & "") & ", " & Val(rsDetails("Mada") & "") & ", " & Val(rsDetails("Visa") & "") & ", " & Val(rsDetails("Deferred") & "") & ", " & _
              Val(rsDetails("AmountH") & "") & ", " & Val(rsDetails("AmountHComm") & "") & ", '" & rsDetails("DetailsPump") & "', '" & _
              Trim(rsDetails("Account_CodeComm") & "") & "', '" & rsDetails("Account_Code") & "', " & Val(rsDetails("IsOther") & "") & ")"

        '  ‰›Ì– «·«” ⁄·«„
                                mLastStep = "Insert missing items into remote server"
mLastSQL = sql
        POSConnection.Execute sql
        If Err.Number <> 0 Then
            frmPopup.ShowMessage "Error in Transaction_Details: " & Err.Description
            Err.Clear
        End If
        rsDetails.MoveNext
    Loop
    rsDetails.Close

    ' «·«‰ ﬁ«· ≈·Ï «·”Ã· «· «·Ì
    rs.MoveNext
Loop
rs.Close
lblWait.Caption = " „ ‰ﬁ· ”‰œ«  «· ÕÊÌ·«  «·„Œ“‰Ì… »‰Ã«Õ"
POSConnection.CommitTrans
End Sub
Private Function SqlStr(v As Variant) As String
    If IsNull(v) Or Trim$(CStr(v & "")) = "" Then
        SqlStr = "NULL"
    Else
        SqlStr = "'" & Replace$(CStr(v), "'", "''") & "'"
    End If
End Function

Private Function SqlNum(v As Variant) As String
    If IsNull(v) Or Trim$(CStr(v & "")) = "" Then
        SqlNum = "NULL"
    Else
        SqlNum = CStr(Val(v))
    End If
End Function

Private Function SqlBit(v As Variant) As String
    If IsNull(v) Or Trim$(CStr(v & "")) = "" Then
        SqlBit = "NULL"
    Else
        SqlBit = IIf(CBool(v), "1", "0")
    End If
End Function

Private Function SqlDateTime(v As Variant) As String
    If IsNull(v) Or Trim$(CStr(v & "")) = "" Then
        SqlDateTime = "NULL"
    Else
        ' yyyy-mm-dd HH:nn:ss
        SqlDateTime = "'" & Format$(CDate(v), "yyyy-mm-dd HH:nn:ss") & "'"
    End If
End Function

'Private Sub Command11_Click()
'
'    Dim sql As String
'    Dim s As String
'    Dim Rs3 As New ADODB.Recordset
'        lblWait.Visible = True
'    '  Õﬁﬁ „‰ ÊÃÊœ « ’«· »«·”Ì—›—
'    If POSlServer.Text = "" Then
'        frmPopup.ShowMessage "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«"
'        Exit Sub
'    End If
'
'    '  ⁄—Ì› «·„ €Ì—«  «·√”«”Ì…
'    Dim mPosD As String
'    Dim mServerD As String
'    mPosD = "[" & POSlServer & "]." & POSDb & ".dbo."
'    mServerD = "[" & SysSQLServerName & "]." & ServerDb & ".dbo."
'
'           mServerD = "[RemoteServer10]." & mDBPOSName & ".dbo."
'
'    lblWait.Visible = True
'
'    ' «· Õﬁﬁ „‰ «·√’‰«› «·„›ﬁÊœ… ›Ì «·”Ì—›— Ê‰ﬁ·Â«
'    On Error GoTo ErrorHandler
'    DoEvents
'    lblWait.Caption = "Ì „ «·«‰ ‰ﬁ· «·«’‰«›"
'    DoEvents
'   ' POSConnection.BeginTrans
'    s = "INSERT INTO " & mServerD & "TblItems (ItemID, ItemCode, ItemName,DefaultSupplier, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, " & _
'        "CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, " & _
'        "ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, barCodeNO, SizeID11) " & _
'        "SELECT T1.ItemID, T1.ItemCode, T1.ItemName,T1.DefaultSupplier, T1.GroupID, T1.HaveSerial, T1.LastUpdate, T1.PurchasePrice, T1.SallingPrice, T1.RequestLimit, " & _
'        "T1.CustomerPrice, T1.HaveGuarantee, T1.GuaranteeValue, T1.GuaranteeType, T1.IsArchive, T1.ItemType, T1.AssbliedItem, T1.RelatedItem, T1.ItemComment, T1.ItemCase, T1.ItemMaking, " & _
'        "T1.ItemMakingNew, T1.code, T1.Branch_NO, T1.Fullcode, T1.prifix, T1.PartNo, T1.CostPrice, T1.ItemNamee, T1.DefaultSupplier, T1.itemSerials, T1.barCodeNO, T1.SizeID11 " & _
'        "FROM " & mPosD & "TblItems T1 " & _
'        "LEFT JOIN " & mServerD & "TblItems T2 ON T1.ItemID = T2.ItemID " & _
'        "WHERE T2.ItemID IS NULL;"
'                                mLastStep = "Insert missing items into remote server"
'mLastSQL = s
'    POSConnection.Execute s
'
'
'
''        s = "INSERT INTO TblItems (ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, " & _
''        "CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, " & _
''        "ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, barCodeNO, SizeID11) " & _
''        "SELECT T1.ItemID, T1.ItemCode, T1.ItemName, T1.GroupID, T1.HaveSerial, T1.LastUpdate, T1.PurchasePrice, T1.SallingPrice, T1.RequestLimit, " & _
''        "T1.CustomerPrice, T1.HaveGuarantee, T1.GuaranteeValue, T1.GuaranteeType, T1.IsArchive, T1.ItemType, T1.AssbliedItem, T1.RelatedItem, T1.ItemComment, T1.ItemCase, T1.ItemMaking, " & _
''        "T1.ItemMakingNew, T1.code, T1.Branch_NO, T1.Fullcode, T1.prifix, T1.PartNo, T1.CostPrice, T1.ItemNamee, T1.DefaultSupplier, T1.itemSerials, T1.barCodeNO, T1.SizeID11 " & _
''        "FROM " & mPosD & "TblItems T1 " & _
''        "LEFT JOIN " & mServerD & "TblItems T2 ON T1.ItemID = T2.ItemID " & _
''        "WHERE T2.ItemID IS NULL;"
''    Cn.Execute s
' '   POSConnection.RollbackTrans
'
' '   POSConnection.CommitTrans
'
'' ‰ﬁ·  ›«’Ì· «·√’‰«› „‰ TblItemsUnits
's = "INSERT INTO " & mServerD & "TblItemsUnits (ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, " & _
'    "MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2) " & _
'    "SELECT T1.ItemID, T1.UnitID, T1.UnitFactor, T1.SecOrder, T1.DefaultUnit, T1.UnitSalesPrice, T1.UnitPurPrice, T1.FactorByDefaultUnit, " & _
'    "T1.MinSelingPrice, T1.ForUnit, T1.MethodCalc, T1.SessionCode, T1.barCodeNo2 " & _
'    "FROM " & mPosD & "TblItemsUnits T1 " & _
'    "LEFT JOIN " & mServerD & "TblItemsUnits T2 ON T1.ItemID = T2.ItemID AND T1.UnitID = T2.UnitID " & _
'    "WHERE T2.ItemID IS NULL;"
'
'                            mLastStep = "Insert missing items into remote server"
'mLastSQL = s
'POSConnection.Execute s
'
'
'
's = "Update T1"
's = s & " SET"
's = s & "     T1.BigUserPw = T2.BigUserPw"
's = s & "     ,T1.BigUserPw2 = T2.BigUserPw2"
's = s & " FROM TblOptions T1"
'
's = s & "  CROSS JOIN " & mServerD & "TblOptions T2 "
'
'                        mLastStep = "Insert missing items into remote server"
'mLastSQL = s
'POSConnection.Execute s
'
'
''    MsgBox " „ ‰ﬁ· «·√’‰«› «·‰«ﬁ’… ≈·Ï «·”Ì—›— »‰Ã«Õ", vbInformation, "‰Ã«Õ"
'    DoEvents
'    lblWait.Caption = " „  ‰ﬁ· «·«’‰«› »‰Ã«Õ"
'    DoEvents
'    lblWait.Visible = True
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "ÕœÀ Œÿ√ √À‰«¡ ⁄„·Ì… «·‰ﬁ·: " & Err.Description, vbCritical, "Œÿ√"
'    lblWait.Visible = False
'    Err.Clear
'
'
'End Sub
'

'Private Sub Command11_Click()
'
'    On Error GoTo ErrorHandler
'
'    Dim sql As String
'    Dim s As String
'    Dim Rs3 As ADODB.Recordset
'
'    Dim mPosD As String
'    Dim mServerD As String
'
'    Set Rs3 = New ADODB.Recordset
'
'    mLastProc = "Command11_Click"
'    mLastSQL = ""
'
'    lblWait.Visible = True
'
'    '  Õﬁﬁ „‰ ÊÃÊœ « ’«· »«·”Ì—›—
'    If POSlServer.Text = "" Then
'        frmPopup.ShowMessage "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«"
'        lblWait.Visible = False
'        Exit Sub
'    End If
'
'    If ConnectionFirst = False Then
'        lblWait.Visible = False
'        Exit Sub
'    End If
'
'    '  ⁄—Ì› «·„ €Ì—«  «·√”«”Ì…
'    mPosD = "[" & POSlServer & "]." & POSDb & ".dbo."
'    mServerD = "[" & SysSQLServerName & "]." & ServerDb & ".dbo."
'    mServerD = "[RemoteServer10]." & mDBPOSName & ".dbo."
'
'    DoEvents
'    lblWait.Caption = "Ì „ «·«‰ ‰ﬁ· «·«’‰«›"
'    DoEvents
'
'    ' 1) ‰ﬁ· «·√’‰«› «·‰«ﬁ’… ≈·Ï «·”Ì—›—
'    s = "INSERT INTO " & mServerD & "TblItems (ItemID, ItemCode, ItemName,DefaultSupplier, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, " & _
'        "CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, " & _
'        "ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, barCodeNO, SizeID11) " & _
'        "SELECT T1.ItemID, T1.ItemCode, T1.ItemName,T1.DefaultSupplier, T1.GroupID, T1.HaveSerial, T1.LastUpdate, T1.PurchasePrice, T1.SallingPrice, T1.RequestLimit, " & _
'        "T1.CustomerPrice, T1.HaveGuarantee, T1.GuaranteeValue, T1.GuaranteeType, T1.IsArchive, T1.ItemType, T1.AssbliedItem, T1.RelatedItem, T1.ItemComment, T1.ItemCase, T1.ItemMaking, " & _
'        "T1.ItemMakingNew, T1.code, T1.Branch_NO, T1.Fullcode, T1.prifix, T1.PartNo, T1.CostPrice, T1.ItemNamee, T1.DefaultSupplier, T1.itemSerials, T1.barCodeNO, T1.SizeID11 " & _
'        "FROM " & mPosD & "TblItems T1 " & _
'        "LEFT JOIN " & mServerD & "TblItems T2 ON T1.ItemID = T2.ItemID " & _
'        "WHERE T2.ItemID IS NULL;"
'
'    If ExecSQL(POSConnection, s, "Command11_Click", "Insert TblItems To Server") = False Then
'        Err.Raise vbObjectError + 1001, , BuildErrMsg(POSConnection, "Command11_Click", s, "›‘· √À‰«¡ ‰ﬁ· TblItems ≈·Ï «·”Ì—›—")
'    End If
'
'    ' 2) ‰ﬁ·  ›«’Ì· «·√’‰«› „‰ TblItemsUnits
'    s = "INSERT INTO " & mServerD & "TblItemsUnits (ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, " & _
'        "MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2) " & _
'        "SELECT T1.ItemID, T1.UnitID, T1.UnitFactor, T1.SecOrder, T1.DefaultUnit, T1.UnitSalesPrice, T1.UnitPurPrice, T1.FactorByDefaultUnit, " & _
'        "T1.MinSelingPrice, T1.ForUnit, T1.MethodCalc, T1.SessionCode, T1.barCodeNo2 " & _
'        "FROM " & mPosD & "TblItemsUnits T1 " & _
'        "LEFT JOIN " & mServerD & "TblItemsUnits T2 ON T1.ItemID = T2.ItemID AND T1.UnitID = T2.UnitID " & _
'        "WHERE T2.ItemID IS NULL;"
'
'    If ExecSQL(POSConnection, s, "Command11_Click", "Insert TblItemsUnits To Server") = False Then
'        Err.Raise vbObjectError + 1002, , BuildErrMsg(POSConnection, "Command11_Click", s, "›‘· √À‰«¡ ‰ﬁ· TblItemsUnits ≈·Ï «·”Ì—›—")
'    End If
'
'    ' 3)  ÕœÌÀ «·»«”Ê—œ«  ›Ì TblOptions
'    s = "Update T1 " & _
'        "SET T1.BigUserPw = T2.BigUserPw, " & _
'        "    T1.BigUserPw2 = T2.BigUserPw2 " & _
'        "FROM TblOptions T1 " & _
'        "CROSS JOIN " & mServerD & "TblOptions T2 "
'
'    If ExecSQL(POSConnection, s, "Command11_Click", "Update TblOptions Passwords") = False Then
'        Err.Raise vbObjectError + 1003, , BuildErrMsg(POSConnection, "Command11_Click", s, "›‘· √À‰«¡  ÕœÌÀ TblOptions")
'    End If
'
'    DoEvents
'    lblWait.Caption = " „ ‰ﬁ· «·«’‰«› »‰Ã«Õ"
'    DoEvents
'    lblWait.Visible = True
'
'    SafeCloseRS Rs3
'    Exit Sub
'
'ErrorHandler:
'    LogAdoErrors POSConnection, "Command11_Click", s, "Œÿ√ ⁄«„ ›Ì Command11_Click"
'
'    MsgBox BuildErrMsg(POSConnection, "Command11_Click", s, "ÕœÀ Œÿ√ √À‰«¡ ⁄„·Ì… «·‰ﬁ·"), vbCritical, "Œÿ√"
'
'    lblWait.Visible = False
'    SafeCloseRS Rs3
'    Err.Clear
'
'End Sub


Private Sub Command11_Click()

    On Error GoTo ErrorHandler

    Dim s As String
    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset

    Dim ItemID As Variant
    Dim UnitID As Variant

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    mLastProc = "Command11_Click"
    mLastSQL = ""

    lblWait.Visible = True
    lblWait.Caption = "Ã«—Ì ‰ﬁ· «·√’‰«›"
    DoEvents

    If Trim$(POSlServer.Text) = "" Then
        frmPopup.ShowMessage "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«"
        lblWait.Visible = False
        Exit Sub
    End If

    If ConnectionFirst = False Then
        lblWait.Visible = False
        Exit Sub
    End If

    If Trim$(POSDb & "") = "" Then
        MsgBox "«”„ ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ… €Ì— „Õœœ", vbCritical, "Œÿ√"
        lblWait.Visible = False
        Exit Sub
    End If

    If Trim$(ServerDb & "") = "" Then
        MsgBox "«”„ ﬁ«⁄œ… »Ì«‰«  «·”Ì—›— €Ì— „Õœœ", vbCritical, "Œÿ√"
        lblWait.Visible = False
        Exit Sub
    End If

    '========================
    ' 1) ‰ﬁ· «·√’‰«› «·‰«ﬁ’… „‰ «·”Ì—›— ≈·Ï «·‰ﬁÿ…
    ' ﬁ—«¡… „‰ Cn - ﬂ «»… ⁄·Ï POSConnection
    '========================
    lblWait.Caption = "Ã«—Ì ‰ﬁ· «·√’‰«› „‰ «·”Ì—›— ≈·Ï «·‰ﬁÿ…"
    DoEvents

    s = ""
    s = s & "SELECT "
    s = s & "ItemID, ItemCode, ItemName, DefaultSupplier, GroupID, HaveSerial, LastUpdate, "
    s = s & "PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, HaveGuarantee, "
    s = s & "GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, "
    s = s & "ItemComment, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode, "
    s = s & "prifix, PartNo, CostPrice, ItemNamee, itemSerials, barCodeNO, SizeID11 "
    s = s & "FROM TblItems "
    s = s & "ORDER BY ItemID"

    mLastSQL = s
    rsSrc.Open s, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsSrc.EOF

        ItemID = rsSrc("ItemID").Value

        SafeCloseRS rsChk
        s = "SELECT ItemID FROM TblItems WHERE ItemID = " & Val(ItemID & "")
        mLastSQL = s
        Set rsChk = New ADODB.Recordset

        rsChk.Open s, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            SafeCloseRS rsChk

            s = ""
            s = s & "INSERT INTO TblItems ("
            s = s & "ItemID, ItemCode, ItemName, DefaultSupplier, GroupID, HaveSerial, LastUpdate, "
            s = s & "PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, HaveGuarantee, "
            s = s & "GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, "
            s = s & "ItemComment, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode, "
            s = s & "prifix, PartNo, CostPrice, ItemNamee, itemSerials, barCodeNO, SizeID11"
            s = s & ") VALUES ("
            s = s & SqlNum(rsSrc("ItemID").Value) & ","
            s = s & SqlStr(rsSrc("ItemCode").Value) & ","
            s = s & SqlStr(rsSrc("ItemName").Value) & ","
            s = s & SqlNum(rsSrc("DefaultSupplier").Value) & ","
            s = s & SqlNum(rsSrc("GroupID").Value) & ","
            s = s & SqlNum(rsSrc("HaveSerial").Value) & ","
            s = s & SqlDateTime(rsSrc("LastUpdate").Value) & ","
            s = s & SqlNum(rsSrc("PurchasePrice").Value) & ","
            s = s & SqlNum(rsSrc("SallingPrice").Value) & ","
            s = s & SqlNum(rsSrc("RequestLimit").Value) & ","
            s = s & SqlNum(rsSrc("CustomerPrice").Value) & ","
            s = s & SqlNum(rsSrc("HaveGuarantee").Value) & ","
            s = s & SqlNum(rsSrc("GuaranteeValue").Value) & ","
            s = s & SqlNum(rsSrc("GuaranteeType").Value) & ","
            s = s & SqlNum(rsSrc("IsArchive").Value) & ","
            s = s & SqlNum(rsSrc("ItemType").Value) & ","
            s = s & SqlNum(rsSrc("AssbliedItem").Value) & ","
            s = s & SqlNum(rsSrc("RelatedItem").Value) & ","
            s = s & SqlStr(rsSrc("ItemComment").Value) & ","
            s = s & SqlNum(rsSrc("ItemCase").Value) & ","
            s = s & SqlNum(rsSrc("ItemMaking").Value) & ","
            s = s & SqlNum(rsSrc("ItemMakingNew").Value) & ","
            s = s & SqlStr(rsSrc("code").Value) & ","
            s = s & SqlNum(rsSrc("Branch_NO").Value) & ","
            s = s & SqlStr(rsSrc("Fullcode").Value) & ","
            s = s & SqlStr(rsSrc("prifix").Value) & ","
            s = s & SqlStr(rsSrc("PartNo").Value) & ","
            s = s & SqlNum(rsSrc("CostPrice").Value) & ","
            s = s & SqlStr(rsSrc("ItemNamee").Value) & ","
            s = s & SqlNum(rsSrc("itemSerials").Value) & ","
            s = s & SqlStr(rsSrc("barCodeNO").Value) & ","
            s = s & SqlNum(rsSrc("SizeID11").Value)
            s = s & ")"

            If ExecSQL(POSConnection, s, "Command11_Click", "Insert TblItems To POS") = False Then
                Err.Raise vbObjectError + 1001, , BuildErrMsg(POSConnection, "Command11_Click", s, "›‘· √À‰«¡ ‰ﬁ· TblItems ≈·Ï «·‰ﬁÿ…")
            End If
        Else
            SafeCloseRS rsChk
        End If

        rsSrc.MoveNext
    Loop
    SafeCloseRS rsSrc

    '========================
    ' 2) ‰ﬁ· ÊÕœ«  «·√’‰«› «·‰«ﬁ’… „‰ «·”Ì—›— ≈·Ï «·‰ﬁÿ…
    '========================
    lblWait.Caption = "Ã«—Ì ‰ﬁ· ÊÕœ«  «·√’‰«› „‰ «·”Ì—›— ≈·Ï «·‰ﬁÿ…"
    DoEvents

    s = ""
    s = s & "SELECT "
    s = s & "ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, "
    s = s & "FactorByDefaultUnit, MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2 "
    s = s & "FROM TblItemsUnits "
    s = s & "ORDER BY ItemID, UnitID"

    mLastSQL = s
    Set rsSrc = New ADODB.Recordset
    rsSrc.Open s, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsSrc.EOF

        ItemID = rsSrc("ItemID").Value
        UnitID = rsSrc("UnitID").Value

        SafeCloseRS rsChk
        s = "SELECT ItemID, UnitID FROM TblItemsUnits WHERE ItemID = " & Val(ItemID & "") & " AND UnitID = " & Val(UnitID & "")
        mLastSQL = s
            Set rsChk = New ADODB.Recordset
        rsChk.Open s, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            SafeCloseRS rsChk

            s = ""
            s = s & "INSERT INTO TblItemsUnits ("
            s = s & "ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, "
            s = s & "FactorByDefaultUnit, MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2"
            s = s & ") VALUES ("
            s = s & SqlNum(rsSrc("ItemID").Value) & ","
            s = s & SqlNum(rsSrc("UnitID").Value) & ","
            s = s & SqlNum(rsSrc("UnitFactor").Value) & ","
            s = s & SqlNum(rsSrc("SecOrder").Value) & ","
            s = s & SqlNum(rsSrc("DefaultUnit").Value) & ","
            s = s & SqlNum(rsSrc("UnitSalesPrice").Value) & ","
            s = s & SqlNum(rsSrc("UnitPurPrice").Value) & ","
            s = s & SqlNum(rsSrc("FactorByDefaultUnit").Value) & ","
            s = s & SqlNum(rsSrc("MinSelingPrice").Value) & ","
            s = s & SqlNum(rsSrc("ForUnit").Value) & ","
            s = s & SqlNum(rsSrc("MethodCalc").Value) & ","
            s = s & SqlStr(rsSrc("SessionCode").Value) & ","
            s = s & SqlStr(rsSrc("barCodeNo2").Value)
            s = s & ")"

            If ExecSQL(POSConnection, s, "Command11_Click", "Insert TblItemsUnits To POS") = False Then
                Err.Raise vbObjectError + 1002, , BuildErrMsg(POSConnection, "Command11_Click", s, "›‘· √À‰«¡ ‰ﬁ· TblItemsUnits ≈·Ï «·‰ﬁÿ…")
            End If
        Else
            SafeCloseRS rsChk
        End If

        rsSrc.MoveNext
    Loop
    SafeCloseRS rsSrc

    '========================
    ' 3)  ÕœÌÀ »«”Ê—œ«  TblOptions ⁄·Ï «·‰ﬁÿ… „‰ «·”Ì—›—
    '========================
    lblWait.Caption = "Ã«—Ì  ÕœÌÀ ≈⁄œ«œ«  «·‰ﬁÿ…"
    DoEvents

    s = "SELECT TOP 1 BigUserPw, BigUserPw2 FROM TblOptions"
    mLastSQL = s
     Set rsSrc = New ADODB.Recordset

    rsSrc.Open s, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsSrc.EOF Then
        s = ""
        s = s & "UPDATE TblOptions "
        s = s & "SET BigUserPw = " & SqlStr(rsSrc("BigUserPw").Value) & ", "
        s = s & "    BigUserPw2 = " & SqlStr(rsSrc("BigUserPw2").Value)

        If ExecSQL(POSConnection, s, "Command11_Click", "Update TblOptions Passwords On POS") = False Then
            Err.Raise vbObjectError + 1003, , BuildErrMsg(POSConnection, "Command11_Click", s, "›‘· √À‰«¡  ÕœÌÀ TblOptions ⁄·Ï «·‰ﬁÿ…")
        End If
    End If
    SafeCloseRS rsSrc

    lblWait.Caption = " „ ‰ﬁ· «·√’‰«› »‰Ã«Õ"
    lblWait.Visible = True
    DoEvents

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
    Exit Sub

ErrorHandler:
    LogAdoErrors POSConnection, "Command11_Click", s, "Œÿ√ ⁄«„ ›Ì Command11_Click"
    MsgBox BuildErrMsg(POSConnection, "Command11_Click", s, "ÕœÀ Œÿ√ √À‰«¡ ⁄„·Ì… «·‰ﬁ·"), vbCritical, "Œÿ√"

    lblWait.Visible = False
    SafeCloseRS rsChk
    SafeCloseRS rsSrc
    Err.Clear

End Sub

Private Sub Command16_Click()

End Sub
Private Sub cmdTestInsert_Click()

    On Error GoTo errHandler

    Dim s As String
    Dim rs As ADODB.Recordset

    If ConnectionFirst = False Then Exit Sub

    Set rs = New ADODB.Recordset

    s = ""
    s = s & "INSERT INTO DelMe ("
    s = s & "Code, Name, Account_Serial, Account_Code, CusId"
    s = s & ") VALUES ("
    s = s & "'" & Replace("T" & Format$(Now, "yyyymmddhhnnss"), "'", "''") & "',"
    s = s & "N'" & Replace("Test Insert " & Format$(Now, "yyyy-mm-dd hh:nn:ss"), "'", "''") & "',"
    s = s & "1,"
    s = s & "'" & Replace("ACC-TEST", "'", "''") & "',"
    s = s & "1"
    s = s & ")"

    Cn.Execute s, , adExecuteNoRecords

    MsgBox "Insert OK", vbInformation

    Exit Sub

errHandler:
    MsgBox "Insert Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & s, vbCritical

End Sub

'
Private Function ForceTestMode(Optional ByVal ShowMsg As Boolean = True) As Boolean

    On Error GoTo errHandler

    ForceTestMode = False

    ' ·«“„ «·‰ﬁÿ…  ﬂÊ‰ „ Œ «—…
    If Trim$(POSlServer.Text & "") = "" Then
        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
        Exit Function
    End If

    ' ·«“„ «”„ «·”Ì—›— «·„—ﬂ“Ì ÌﬂÊ‰ „ÊÃÊœ
    If Trim$(SysSQLServerName & "") = "" Then
        MsgBox "«”„ «·”Ì—›— «·„—ﬂ“Ì €Ì— „Õœœ", vbCritical, "Œÿ√"
        Exit Function
    End If

    ' ≈Ã»«— «· ‘€Ì· ⁄·Ï Test
    TxtServerDataBaseName.Text = "Test"
    TxtPOSDB.Text = "Test"

    ServerDb = "Test"
    POSDb = "Test"

    If ShowMsg Then
        MsgBox "DEBUG TEST MODE" & vbCrLf & vbCrLf & _
               "Central Server = " & SysSQLServerName & vbCrLf & _
               "Central DB = " & ServerDb & vbCrLf & vbCrLf & _
               "POS Server = " & POSlServer.Text & vbCrLf & _
               "POS DB = " & POSDb, _
               vbInformation, " ‘€Ì· ⁄·Ï ﬁÊ«⁄œ Test"
    End If

    If ConnectionFirst = False Then
        Exit Function
    End If

    '  √ﬂÌœ ‰Â«∆Ì
    If UCase$(Trim$(ServerDb & "")) <> "TEST" Then
        MsgBox "›‘· ÷»ÿ ﬁ«⁄œ… «·”Ì—›— ⁄·Ï Test", vbCritical, "Œÿ√"
        Exit Function
    End If

    If UCase$(Trim$(POSDb & "")) <> "TEST" Then
        MsgBox "›‘· ÷»ÿ ﬁ«⁄œ… «·‰ﬁÿ… ⁄·Ï Test", vbCritical, "Œÿ√"
        Exit Function
    End If

    ForceTestMode = True
    Exit Function

errHandler:
    MsgBox "ForceTestMode Error: " & Err.Description, vbCritical, "Error"
    ForceTestMode = False

End Function
Private Sub Command17_Click()

End Sub



'
Private Sub Command14_Debug_Click()

    On Error GoTo ErrorHandler

    '========================
    ' Declarations
    '========================
    Dim POSCn As ADODB.Connection
    Dim rsCnt As ADODB.Recordset
    Dim rsTrans As ADODB.Recordset
    Dim rsDetails As ADODB.Recordset
    Dim rsValueAdded As ADODB.Recordset
    Dim rsPayments As ADODB.Recordset
    Dim rsSer As ADODB.Recordset

    Dim BatchSize As Integer
    Dim recCounter As Long
    Dim recCount As Long
    Dim BatchNo As Long
    Dim TotalInvoices As Long

    Dim transBatchSQL As String
    Dim detailsBatchSQL As String
    Dim valueAddedBatchSQL As String
    Dim paymentsBatchSQL2 As String
    Dim paymentsBatchSQL As String
    Dim LastSQL As String
    Dim sql As String
    Dim transSQL As String
    Dim detailSQL As String
    Dim valueSQL As String
    Dim paymentSQL As String
    Dim paymentSQL2 As String

    Dim SessionCode As String
    Dim CurrentInvoiceNo As String

    Dim mTimeStart As Date
    Dim elapsedSec As Long, elapsedMin As Long

    Dim direction As String, kind As String
    Dim FetchSize As Long
    Dim mServerD As String

    Dim CarOilChangeDate As Date, RecTime As Date, mTimeIn As String
    Dim ActualDeliveryDate As Date, LatestDeliveryDate As Date
    Dim FromTransaction_Date As Date
    Dim FromTransaction_Type As Long
    Dim FromTransaction_ID As Double

    Dim PayMentType As Long, cusID As Long, BranchID As Integer, BoxID As Long, BillBasedOn As Double
    Dim VAT As Double, VATYou As Double, NoteId As Long, Trans_DiscountType As Long
    Dim Trans_Discount As Double, TaxValue As Double, order_no As Long, SaleType As Long
    Dim TaxAddValue As Double, NetValue As Double, Transaction_NetValue As Double, DepandToConv As Long
    Dim CarTypeID As Long, OilsTypesID As Long, YearFact As Long, FixesAssetsID As Long, ColorID2 As Long
    Dim KM As Double, PPointID As Long, SupplerID As Long, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double
    Dim CarEnginoil As Double, CarGearOil As Double, InvoiceTypeCodeID As Long
    Dim storeID As Variant, userID As Variant, Emp_ID As Variant
    Dim NoteSerial As String, NoteSerial1 As String, TransactionComment As String
    Dim CashCustomerName As String, CashCustomerPhone As String
    Dim cleanCashCustomerName As String
    Dim PlateNo As String, Shaseh As String, CarMeter As String
    Dim Chasee As String, Phone2 As String
    Dim CIBAN As String
    Dim InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String
    Dim paymentnote As String, PaymentMeansCode As String

    ' Counters for reconcile
    Dim SrcHeads As Long, SrcDet As Long, SrcVAT As Long, SrcPay As Long, SrcPay2 As Long
    Dim DstHeads As Long, DstDet As Long, DstVAT As Long, DstPay As Long, DstPay2 As Long
    Dim srcCount As Long, dstCount As Long

    ' Checksums
    Dim srcQty As Double, dstQty As Double
    Dim SrcAmount As Currency, DstAmount As Currency
    Dim SrcVATSum As Currency, DstVATSum As Currency
    Dim SrcTPay As Currency, DstTPay As Currency
    Dim SrcSPay As Currency, DstSPay As Currency
    Dim epsQty As Double
    Dim epsMoney As Currency

    ' Misc
    Dim currentDestTransactionID As String
    Dim tmpRecTime As Variant
    Dim v As Variant
    Dim r As Integer
    Dim DebugLogFile As String
    Dim Recorddate As Date

    '========================
    ' Initial values
    '========================
    BatchSize = 50
    recCounter = 0
    recCount = 0
    BatchNo = 0
    FetchSize = 0

    transBatchSQL = ""
    detailsBatchSQL = ""
    valueAddedBatchSQL = ""
    paymentsBatchSQL2 = ""
    paymentsBatchSQL = ""
    LastSQL = ""

    direction = "POS->Server"
    kind = "Sales-DEBUG"

    mTimeStart = Now
    txtStartTime = mTimeStart

    '========================
    ' Validation
    '========================
    If Trim$(POSlServer.Text) = "" Then
        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
        Exit Sub
    End If

    If PrepareDebugTestConnections = False Then Exit Sub
    WritePhaseLog "DEBUG Connection Info", _
        "ServerDb=" & ServerDb & _
        " | POSDb=" & POSDb & _
        " | SysSQLServerName=" & SysSQLServerName & _
        " | POSServer(Text)=" & Trim$(POSlServer.Text) & _
        " | POSServer(Var)=" & Trim$(POSServer)

    If Not Cn Is Nothing Then
        WritePhaseLog "DEBUG Cn State", "Cn.State=" & CStr(Cn.State)
        WritePhaseLog "DEBUG Cn.ConnectionString", Cn.ConnectionString
    Else
        WritePhaseLog "DEBUG Cn State", "Cn Is Nothing"
    End If

    If Not POSConnection Is Nothing Then
        WritePhaseLog "DEBUG POSConnection State", "POSConnection.State=" & CStr(POSConnection.State)
        WritePhaseLog "DEBUG POSConnection.ConnectionString", POSConnection.ConnectionString
    Else
        WritePhaseLog "DEBUG POSConnection State", "POSConnection Is Nothing"
    End If
    
    lblWait.Visible = True
    lblWait.Caption = "DEBUG MODE - Ã«—Ì »œ¡ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄« ..."
    DoEvents
    MousePointer = vbHourglass

    SessionCode = Format$(Now, "yyyymmddhhnnss")
    DebugLogFile = GetDebugLogFileName(SessionCode)

    DebugWriteLine DebugLogFile, "========== START DEBUG COMMAND14 =========="
    DebugWriteLine DebugLogFile, "ServerDb=" & ServerDb
    DebugWriteLine DebugLogFile, "POSDb=" & POSDb
    DebugWriteLine DebugLogFile, "POSServer=" & POSlServer.Text
    DebugWriteLine DebugLogFile, "SessionCode=" & SessionCode

    UpdateTransferCaption "DEBUG - Ã«—Ì  ÃÂÌ“ Ã·”… «·‰ﬁ·", 0, 0, SessionCode, mTimeStart

    '========================
    ' Open local POS connection on TEST
    '========================
    Set POSCn = New ADODB.Connection
    POSCn.CursorLocation = adUseServer
    POSCn.ConnectionTimeout = 5000
    POSCn.CommandTimeout = 5000
    POSCn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
                             ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                             ";Initial Catalog=Test;Data Source=" & POSlServer.Text
    POSCn.Open
    If Not POSCn Is Nothing Then
        WritePhaseLog "DEBUG POSCn State", "POSCn.State=" & CStr(POSCn.State)
        WritePhaseLog "DEBUG POSCn.ConnectionString", POSCn.ConnectionString
    Else
        WritePhaseLog "DEBUG POSCn State", "POSCn Is Nothing"
    End If

    ' Central connection already prepared by ConnectionFirst -> TEST
    Cn.CursorLocation = adUseServer
    mServerD = "dbo."

    DebugWriteLine DebugLogFile, "Opened POSCn on Test successfully"
    DebugWriteLine DebugLogFile, "Opened Cn on Test successfully"

    Text3 = "Query: " & GetQuery

    '========================
    ' Tag source transactions
    '========================
    LastSQL = "UPDATE Transactions SET SessionCode = '" & SessionCode & "' WHERE IsNull(Copied,0)=0 AND " & GetQuery
    DebugWriteSQL DebugLogFile, "Tag source transactions", LastSQL
    POSCn.Execute LastSQL

    UpdateTransferCaption "DEBUG -  „  ⁄·Ì„ «·›Ê« Ì— «·„—«œ ‰ﬁ·Â«", 0, 0, SessionCode, mTimeStart

    '========================
    ' Get total count first
    '========================
    LastSQL = "SELECT COUNT(*) AS Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND " & GetQuery
    DebugWriteSQL DebugLogFile, "Count tagged transactions", LastSQL

    Set rsCnt = POSCn.Execute(LastSQL)
    TotalInvoices = CLng(rsCnt!Cnt)
    rsCnt.Close
    Set rsCnt = Nothing

    DebugWriteLine DebugLogFile, "TotalInvoices=" & CStr(TotalInvoices)

    UpdateTransferCaption "DEBUG -  „ «·⁄ÀÊ— ⁄·Ï ›Ê« Ì— ··‰ﬁ·", 0, TotalInvoices, SessionCode, mTimeStart

    If TotalInvoices = 0 Then
        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
        GoTo EndSub
    End If

    '========================
    ' Open source transactions
    '========================
    Set rsTrans = New ADODB.Recordset
    LastSQL = "SELECT * FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND " & GetQuery & " ORDER BY Transaction_ID"
    DebugWriteSQL DebugLogFile, "Open source transactions", LastSQL

    rsTrans.Open LastSQL, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rsTrans.EOF Then
        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
        GoTo EndSub
    End If

    '========================
    ' Start SQL transaction on central server
    '========================
    LastSQL = "SET XACT_ABORT ON;"
    DebugWriteSQL DebugLogFile, "SET XACT_ABORT ON", LastSQL
    Cn.Execute LastSQL
    Cn.BeginTrans
    DebugWriteLine DebugLogFile, "Transaction started on server"

    '========================
    ' Main Loop
    '========================
    Do While Not rsTrans.EOF

        If Trim$(rsTrans("NoteSerial1").Value & "") <> "" Then
            CurrentInvoiceNo = Trim$(rsTrans("NoteSerial1").Value & "")
        Else
            CurrentInvoiceNo = CStr(Val(rsTrans("Transaction_ID").Value & ""))
        End If

        If recCounter = 0 Or recCounter Mod 5 = 0 Then
            UpdateTransferCaption "DEBUG - Ã«—Ì ﬁ—«¡… «·›« Ê—… —ﬁ„ " & CurrentInvoiceNo, recCounter + 1, TotalInvoices, SessionCode, mTimeStart
        End If

        recCount = recCount + 1
        DebugWriteLine DebugLogFile, "Reading invoice SrcTransactionID=" & CStr(Val(rsTrans("Transaction_ID").Value & "")) & ", NoteSerial1=" & CurrentInvoiceNo

        '========================
        ' Read all needed values
        '========================
        PayMentType = Val(rsTrans("PaymentType").Value & "")
        FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
        cusID = Val(rsTrans("CusID").Value & "")
        storeID = Val(rsTrans("StoreID").Value & "")
        userID = Val(rsTrans("UserID").Value & "")
        Emp_ID = Val(rsTrans("Emp_ID").Value & "")
        BranchID = Val(rsTrans("BranchID").Value & "")
        BoxID = Val(rsTrans("BoxID").Value & "")
        BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
        mTimeIn = Trim$(rsTrans("TimeIn").Value & "")
        VAT = Val(rsTrans("VAT").Value & "")
        VATYou = Val(rsTrans("VATYou").Value & "")
        NoteSerial = rsTrans("NoteSerial").Value & ""
        NoteSerial1 = rsTrans("NoteSerial1").Value & ""
        NoteId = Val(rsTrans("NoteId").Value & "")
        TransactionComment = rsTrans("TransactionComment").Value & ""
        Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
        FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
        Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
        TaxValue = Val(rsTrans("TaxValue").Value & "")
        order_no = Val(rsTrans("order_no").Value & "")
        SaleType = Val(rsTrans("SaleType").Value & "")
        CashCustomerName = rsTrans("CashCustomerName").Value & ""
        cleanCashCustomerName = CashCustomerName
        TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
        CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
        NetValue = Val(rsTrans("NetValue").Value & "")
        Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
        DepandToConv = Val(rsTrans("DepandToConv").Value & "")
        CarTypeID = Val(rsTrans("CarTypeID").Value & "")
        PlateNo = rsTrans("PlateNo").Value & ""
        OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
        YearFact = Val(rsTrans("YearFact").Value & "")
        Shaseh = rsTrans("Shaseh").Value & ""
        CarMeter = rsTrans("CarMeter").Value & ""
        FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
        ColorID2 = Val(rsTrans("ColorID2").Value & "")
        KM = Val(rsTrans("KM").Value & "")
        Chasee = rsTrans("Chasee").Value & ""
        PPointID = Val(rsTrans("PPointID").Value & "")
        Phone2 = rsTrans("Phone2").Value & ""
        SupplerID = Val(rsTrans("SupplerID").Value & "")
        Ser = Val(rsTrans("Ser").Value & "")
        CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
        CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
        CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
        CarGearOil = Val(rsTrans("CarGearOil").Value & "")

        If Trim$(rsTrans("CarOilChangeDate").Value & "") = "" Then
            CarOilChangeDate = Date
        Else
            CarOilChangeDate = rsTrans("CarOilChangeDate").Value
        End If

        CIBAN = rsTrans("CIBAN").Value & ""

        tmpRecTime = rsTrans("RecTime").Value
        v = tmpRecTime
        If IsDate(v) Then
            If Year(CDate(v)) = 1899 And Month(CDate(v)) = 12 And Day(CDate(v)) = 30 Then
                RecTime = Time
            Else
                RecTime = CDate(v)
            End If
        Else
            RecTime = Time
        End If

        If Trim$(rsTrans("ActualDeliveryDate").Value & "") = "" Then
            ActualDeliveryDate = Date
        Else
            ActualDeliveryDate = rsTrans("ActualDeliveryDate").Value
        End If

        If Trim$(rsTrans("LatestDeliveryDate").Value & "") = "" Then
            LatestDeliveryDate = ActualDeliveryDate
        Else
            LatestDeliveryDate = rsTrans("LatestDeliveryDate").Value
        End If

        InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
        InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
        DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
        TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
        paymentnote = rsTrans("paymentnote").Value & ""
        PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""
        FromTransaction_Date = rsTrans("Transaction_Date").Value

        If Val(rsTrans("POSBillType").Value & "") = 0 Then
            NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
            NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
        End If

        TransactionComment = " ›« Ê—… „‰ﬁÊ·… „‰ " & POSname.Text & "   " & _
                             "   —ﬁ„ «·›« Ê—… " & NoteSerial1

        '========================
        ' Reserve destination Transaction_ID
        '========================
        LastSQL = "EXEC dbo.ReserveTransactionId"
        DebugWriteSQL DebugLogFile, "ReserveTransactionId", LastSQL

        Set rsSer = Cn.Execute(LastSQL)
        If Not rsSer.EOF Then
            currentDestTransactionID = CStr(rsSer.Fields("NewId").Value)
        Else
            Err.Raise vbObjectError + 500, , "·„ Ì „ ≈—Ã«⁄ Transaction_ID ÃœÌœ „‰ «·”Ì—›—"
        End If
        rsSer.Close
        Set rsSer = Nothing

        DebugWriteLine DebugLogFile, "Reserved DestTransactionID=" & currentDestTransactionID & " for SrcTransactionID=" & CStr(FromTransaction_ID)

        '========================
        ' Build Transactions INSERT
        '========================
        transSQL = "INSERT INTO " & mServerD & "Transactions (" & _
        "Transaction_ID, Transaction_Date,TimeIn ,TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
        "CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, NoteSerial1, " & _
        "NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, OldNoteserial, OldNoteId, " & _
        "OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, SaleType, CashCustomerName, " & _
        "TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, DepandToConv, CarTypeID, " & _
        "PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, ColorID2, KM, Chasee, PPointID, Phone2, " & _
        "SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, CarGearOil, CarOilChangeDate, CIBAN, RecTime, " & _
        "ActualDeliveryDate, LatestDeliveryDate, InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, " & _
        "TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("

        transSQL = transSQL & currentDestTransactionID & "," & SQLDate(FromTransaction_Date, True) & ",'" & Trim$(mTimeIn) & "'," & _
        Val(rsTrans("TypeInvoice").Value & "") & ",'" & Replace(rsTrans("Transaction_Serial").Value & "", "'", "''") & "'," & _
        FromTransaction_Type & "," & PayMentType & "," & cusID & "," & storeID & "," & userID & "," & _
        Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & ",'" & _
        Replace(NoteSerial, "'", "''") & "','" & Replace(NoteSerial1, "'", "''") & "'," & NoteId & ",1,'" & Replace(TransactionComment, "'", "''") & "','" & _
        SessionCode & "'," & IIf(Val(rsTrans("POSBillType").Value & "") = 0, 1, Val(rsTrans("POSBillType").Value & "")) & ",'" & Replace(rsTrans("NoteSerial1").Value & "", "'", "''") & "','" & Replace(Trim$(rsTrans("NoteSerial").Value & ""), "'", "''") & "'," & _
        Val(rsTrans("NoteId").Value & "") & "," & Val(rsTrans("Transaction_ID").Value & "") & "," & Trans_DiscountType & "," & Val(Trans_Discount & "") & "," & _
        TaxValue & ",'" & order_no & "'," & SaleType & ",'" & Replace(cleanCashCustomerName, "'", "''") & "'," & _
        TaxAddValue & ",'" & Replace(CashCustomerPhone, "'", "''") & "'," & SQLDate(rsTrans("last_changed").Value, True) & ","

        transSQL = transSQL & NetValue & "," & Transaction_NetValue & "," & IIf(Val(DepandToConv & "") <> 0, 1, 0) & "," & _
        CarTypeID & ",'" & Replace(PlateNo, "'", "''") & "'," & OilsTypesID & "," & YearFact & ",'" & Replace(Shaseh, "'", "''") & "','" & Replace(CarMeter, "'", "''") & "'," & _
        FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Replace(Chasee, "'", "''") & "'," & PPointID & ",'" & Replace(Phone2, "'", "''") & "'," & _
        SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & CarEnginoil & "," & _
        CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & Replace(CIBAN, "'", "''") & "'," & SQLDate(RecTime, True) & ","

        transSQL = transSQL & SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
        InvoiceTypeCodeID & ",'" & Replace(InvoiceTypeCodename, "'", "''") & "','" & Replace(DocumentCurrencyCode, "'", "''") & "','" & _
        Replace(TaxCurrencyCode, "'", "''") & "','" & Replace(paymentnote, "'", "''") & "','" & Replace(PaymentMeansCode, "'", "''") & "')"

        transBatchSQL = transBatchSQL & transSQL & vbCrLf

        '========================
        ' Transaction_Details
        '========================
        sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
        Set rsDetails = New ADODB.Recordset
        rsDetails.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

        Do While Not rsDetails.EOF

            detailSQL = "INSERT INTO " & mServerD & "Transaction_Details (" & _
                "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
                "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
                "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES ("

            detailSQL = detailSQL & currentDestTransactionID & "," & Val(rsDetails("Item_ID").Value & "") & "," & Val(rsDetails("ItemCase").Value & "") & "," & _
                Val(rsDetails("Quantity").Value & "") & "," & Val(rsDetails("Price").Value & "") & "," & Val(rsDetails("ItemDiscountType").Value & "") & "," & _
                Val(rsDetails("ItemDiscount").Value & "") & "," & Val(rsDetails("ShowQty").Value & "") & "," & Val(rsDetails("showPrice").Value & "") & "," & _
                Val(rsDetails("UnitId").Value & "") & "," & Val(rsDetails("ColorID").Value & "") & "," & Val(rsDetails("ItemSize").Value & "") & "," & _
                Val(rsDetails("ClassId").Value & "") & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo").Value & "") & "," & _
                Val(rsDetails("PumpId").Value & "") & "," & Val(rsDetails("PrevQty").Value & "") & ",'" & Replace(Trim$(rsDetails("PrintName").Value & ""), "'", "''") & "'," & _
                Val(rsDetails("Cash").Value & "") & "," & Val(rsDetails("Mada").Value & "") & "," & Val(rsDetails("Visa").Value & "") & "," & _
                Val(rsDetails("Deferred").Value & "") & "," & Val(rsDetails("AmountH").Value & "") & "," & Val(rsDetails("AmountHComm").Value & "") & ","

            detailSQL = detailSQL & "'" & Replace(Trim$(rsDetails("DetailsPump").Value & ""), "'", "''") & "','" & Replace(Trim$(rsDetails("Account_CodeComm").Value & ""), "'", "''") & "','" & _
                Replace(Trim$(rsDetails("Account_Code").Value & ""), "'", "''") & "'," & IIf(Val(rsDetails("IsOther").Value & "") <> 0, 1, 0) & ")"

            detailsBatchSQL = detailsBatchSQL & detailSQL & vbCrLf
            rsDetails.MoveNext
        Loop

        rsDetails.Close
        Set rsDetails = Nothing

        '========================
        ' TransactionValueAdded
        '========================
        sql = "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
        Set rsValueAdded = New ADODB.Recordset
        rsValueAdded.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

        Do While Not rsValueAdded.EOF

            valueSQL = "INSERT INTO " & mServerD & "TransactionValueAdded (" & _
                       "Transaction_ID, ItemID, Vatyo, VAT, Valu, selectd, Transaction_Type, SessionCode) VALUES ("
            valueSQL = valueSQL & currentDestTransactionID & "," & _
                Val(rsValueAdded("ItemID").Value & "") & "," & _
                Val(rsValueAdded("Vatyo").Value & "") & "," & _
                Val(rsValueAdded("Vat").Value & "") & "," & _
                Val(rsValueAdded("Valu").Value & "") & "," & _
                Val(rsValueAdded("selectd").Value & "") & "," & _
                Val(rsValueAdded("Transaction_Type").Value & "") & ",'" & SessionCode & "')"

            valueAddedBatchSQL = valueAddedBatchSQL & valueSQL & vbCrLf
            rsValueAdded.MoveNext
        Loop

        rsValueAdded.Close
        Set rsValueAdded = Nothing

        '========================
        ' Payments
        '========================
        If Val(rsTrans("Transaction_Type").Value & "") = 21 Or Val(rsTrans("Transaction_Type").Value & "") = 9 Then

            Set rsPayments = New ADODB.Recordset
            sql = "SELECT * FROM TblTransactionPayments WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
            rsPayments.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

            Do While Not rsPayments.EOF

                If IsNull(rsPayments("Recorddate").Value) Or Trim$(rsPayments("Recorddate").Value & "") = "" Then
                    Recorddate = Now
                Else
                    Recorddate = rsPayments("Recorddate").Value
                End If

                paymentSQL = "INSERT INTO " & mServerD & "TblTransactionPayments (" & _
                    "Transaction_ID, boxid, Recorddate, PointID, CurrentCashireID, PaymentID, Value, CardNo, Effect, SessionCode) VALUES ("
                paymentSQL = paymentSQL & currentDestTransactionID & "," & _
                    Val(rsPayments("boxid").Value & "") & "," & SQLDate(Recorddate, True) & "," & _
                    Val(rsPayments("PointID").Value & "") & "," & Val(rsPayments("CurrentCashireID").Value & "") & "," & Val(rsPayments("PaymentID").Value & "") & "," & _
                    Val(rsPayments("Value").Value & "") & ",'" & Replace(rsPayments("CardNo").Value & "", "'", "''") & "'," & _
                    Val(rsPayments("Effect").Value & "") & ",'" & SessionCode & "')"

                paymentsBatchSQL = paymentsBatchSQL & paymentSQL & vbCrLf
                rsPayments.MoveNext
            Loop

            rsPayments.Close
            Set rsPayments = Nothing

            Set rsPayments = New ADODB.Recordset
            sql = "SELECT * FROM TblSalesPayment WHERE TransID = " & Val(rsTrans("Transaction_ID").Value & "")
            rsPayments.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

            Do While Not rsPayments.EOF

                paymentSQL2 = "INSERT INTO " & mServerD & "TblSalesPayment (" & _
                    "TransID, PaymentID, Value) VALUES (" & _
                    currentDestTransactionID & "," & Val(rsPayments("PaymentID").Value & "") & "," & Val(rsPayments("Value").Value & "") & ")"

                paymentsBatchSQL2 = paymentsBatchSQL2 & paymentSQL2 & vbCrLf
                rsPayments.MoveNext
            Loop

            rsPayments.Close
            Set rsPayments = Nothing
        End If

        recCounter = recCounter + 1

        '========================
        ' Execute batch every 50
        '========================
        If recCounter Mod BatchSize = 0 Then

            BatchNo = BatchNo + 1
            
            DebugWriteLine DebugLogFile, "Executing BatchNo=" & CStr(BatchNo) & ", RecCounter=" & CStr(recCounter)

            If transBatchSQL <> "" Then
                LastSQL = transBatchSQL
                DebugWriteSQL DebugLogFile, "EXEC BATCH Transactions", transBatchSQL
                Cn.Execute transBatchSQL
            End If

            If detailsBatchSQL <> "" Then
                LastSQL = detailsBatchSQL
                DebugWriteSQL DebugLogFile, "EXEC BATCH Transaction_Details", detailsBatchSQL
                Cn.Execute detailsBatchSQL
            End If

            If valueAddedBatchSQL <> "" Then
                LastSQL = valueAddedBatchSQL
                DebugWriteSQL DebugLogFile, "EXEC BATCH TransactionValueAdded", valueAddedBatchSQL
                Cn.Execute valueAddedBatchSQL
            End If

            If paymentsBatchSQL2 <> "" Then
                LastSQL = paymentsBatchSQL2
                DebugWriteSQL DebugLogFile, "EXEC BATCH TblSalesPayment", paymentsBatchSQL2
                Cn.Execute paymentsBatchSQL2
            End If

            If paymentsBatchSQL <> "" Then
                LastSQL = paymentsBatchSQL
                DebugWriteSQL DebugLogFile, "EXEC BATCH TblTransactionPayments", paymentsBatchSQL
                Cn.Execute paymentsBatchSQL
            End If

            transBatchSQL = ""
            detailsBatchSQL = ""
            valueAddedBatchSQL = ""
            paymentsBatchSQL2 = ""
            paymentsBatchSQL = ""

            DebugWriteLine DebugLogFile, "BatchNo=" & CStr(BatchNo) & " executed successfully"
        End If

        rsTrans.MoveNext
    Loop

    '========================
    ' Final batch
    '========================
    DebugWriteLine DebugLogFile, "Executing final batch"

    If transBatchSQL <> "" Then
        LastSQL = transBatchSQL
        DebugWriteSQL DebugLogFile, "EXEC FINAL Transactions", transBatchSQL
        Cn.Execute transBatchSQL
    End If

    If detailsBatchSQL <> "" Then
        LastSQL = detailsBatchSQL
        DebugWriteSQL DebugLogFile, "EXEC FINAL Transaction_Details", detailsBatchSQL
        Cn.Execute detailsBatchSQL
    End If

    If valueAddedBatchSQL <> "" Then
        LastSQL = valueAddedBatchSQL
        DebugWriteSQL DebugLogFile, "EXEC FINAL TransactionValueAdded", valueAddedBatchSQL
        Cn.Execute valueAddedBatchSQL
    End If

    If paymentsBatchSQL2 <> "" Then
        LastSQL = paymentsBatchSQL2
        DebugWriteSQL DebugLogFile, "EXEC FINAL TblSalesPayment", paymentsBatchSQL2
        Cn.Execute paymentsBatchSQL2
    End If

    If paymentsBatchSQL <> "" Then
        LastSQL = paymentsBatchSQL
        DebugWriteSQL DebugLogFile, "EXEC FINAL TblTransactionPayments", paymentsBatchSQL
        Cn.Execute paymentsBatchSQL
    End If

    transBatchSQL = ""
    detailsBatchSQL = ""
    valueAddedBatchSQL = ""
    paymentsBatchSQL2 = ""
    paymentsBatchSQL = ""

    '========================
    ' Reconcile and checksum
    '========================
    DebugWriteLine DebugLogFile, "Start reconcile/checksum"

    Set rsCnt = POSCn.Execute("SELECT COUNT(*) Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'")
    SrcHeads = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT COUNT(*) Cnt FROM Transaction_Details d WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    SrcDet = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT COUNT(*) Cnt FROM TransactionValueAdded v WHERE v.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    SrcVAT = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT COUNT(*) Cnt FROM TblTransactionPayments p WHERE p.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    SrcPay = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT COUNT(*) Cnt FROM TblSalesPayment s WHERE s.TransID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    SrcPay2 = CLng(rsCnt!Cnt): rsCnt.Close
    Set rsCnt = Nothing

    Set rsCnt = POSCn.Execute("SELECT SUM(CAST(d.Quantity AS float)) AS SumQty, SUM(CAST(d.Quantity * d.Price AS decimal(18,4))) AS SumAmount FROM Transaction_Details d WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumQty) Then srcQty = 0# Else srcQty = CDbl(rsCnt!SumQty)
        If IsNull(rsCnt!SumAmount) Then SrcAmount = CCur(0) Else SrcAmount = CCur(rsCnt!SumAmount)
    End If
    rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT SUM(CAST(v.Valu AS decimal(18,4))) AS SumVAT FROM TransactionValueAdded v WHERE v.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumVAT) Then SrcVATSum = CCur(0) Else SrcVATSum = CCur(rsCnt!SumVAT)
    End If
    rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT SUM(CAST(p.Value AS decimal(18,4))) AS SumPay FROM TblTransactionPayments p WHERE p.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumPay) Then SrcTPay = CCur(0) Else SrcTPay = CCur(rsCnt!SumPay)
    End If
    rsCnt.Close

    Set rsCnt = POSCn.Execute("SELECT SUM(CAST(s.Value AS decimal(18,4))) AS SumPay2 FROM TblSalesPayment s WHERE s.TransID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumPay2) Then SrcSPay = CCur(0) Else SrcSPay = CCur(rsCnt!SumPay2)
    End If
    rsCnt.Close
    Set rsCnt = Nothing

    Set rsCnt = Cn.Execute("SELECT COUNT(*) Cnt FROM dbo.Transactions WHERE SessionCode='" & SessionCode & "'")
    DstHeads = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT COUNT(*) Cnt FROM dbo.Transaction_Details d JOIN dbo.Transactions t ON t.Transaction_ID=d.Transaction_ID WHERE t.SessionCode='" & SessionCode & "'")
    DstDet = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT COUNT(*) Cnt FROM dbo.TransactionValueAdded v JOIN dbo.Transactions t ON t.Transaction_ID=v.Transaction_ID WHERE t.SessionCode='" & SessionCode & "'")
    DstVAT = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT COUNT(*) Cnt FROM dbo.TblTransactionPayments p JOIN dbo.Transactions t ON t.Transaction_ID=p.Transaction_ID WHERE t.SessionCode='" & SessionCode & "'")
    DstPay = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT COUNT(*) Cnt FROM dbo.TblSalesPayment s JOIN dbo.Transactions t ON t.Transaction_ID=s.TransID WHERE t.SessionCode='" & SessionCode & "'")
    DstPay2 = CLng(rsCnt!Cnt): rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT SUM(CAST(d.Quantity AS float)) AS SumQty, SUM(CAST(d.Quantity * d.Price AS decimal(18,4))) AS SumAmount FROM dbo.Transaction_Details d JOIN dbo.Transactions t ON t.Transaction_ID = d.Transaction_ID WHERE t.SessionCode='" & SessionCode & "'")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumQty) Then dstQty = 0# Else dstQty = CDbl(rsCnt!SumQty)
        If IsNull(rsCnt!SumAmount) Then DstAmount = CCur(0) Else DstAmount = CCur(rsCnt!SumAmount)
    End If
    rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT SUM(CAST(v.Valu AS decimal(18,4))) AS SumVAT FROM dbo.TransactionValueAdded v JOIN dbo.Transactions t ON t.Transaction_ID = v.Transaction_ID WHERE t.SessionCode='" & SessionCode & "'")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumVAT) Then DstVATSum = CCur(0) Else DstVATSum = CCur(rsCnt!SumVAT)
    End If
    rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT SUM(CAST(p.Value AS decimal(18,4))) AS SumPay FROM dbo.TblTransactionPayments p JOIN dbo.Transactions t ON t.Transaction_ID = p.Transaction_ID WHERE t.SessionCode='" & SessionCode & "'")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumPay) Then DstTPay = CCur(0) Else DstTPay = CCur(rsCnt!SumPay)
    End If
    rsCnt.Close

    Set rsCnt = Cn.Execute("SELECT SUM(CAST(s.Value AS decimal(18,4))) AS SumPay2 FROM dbo.TblSalesPayment s JOIN dbo.Transactions t ON t.Transaction_ID = s.TransID WHERE t.SessionCode='" & SessionCode & "'")
    If Not rsCnt.EOF Then
        If IsNull(rsCnt!SumPay2) Then DstSPay = CCur(0) Else DstSPay = CCur(rsCnt!SumPay2)
    End If
    rsCnt.Close
    Set rsCnt = Nothing

    epsQty = 0.0001
    epsMoney = 0.01

    DebugWriteLine DebugLogFile, "SrcHeads=" & SrcHeads & ", DstHeads=" & DstHeads
    DebugWriteLine DebugLogFile, "SrcDet=" & SrcDet & ", DstDet=" & DstDet
    DebugWriteLine DebugLogFile, "SrcVAT=" & SrcVAT & ", DstVAT=" & DstVAT
    DebugWriteLine DebugLogFile, "SrcPay=" & SrcPay & ", DstPay=" & DstPay
    DebugWriteLine DebugLogFile, "SrcPay2=" & SrcPay2 & ", DstPay2=" & DstPay2
    DebugWriteLine DebugLogFile, "srcQty=" & srcQty & ", dstQty=" & dstQty
    DebugWriteLine DebugLogFile, "SrcAmount=" & SrcAmount & ", DstAmount=" & DstAmount
    DebugWriteLine DebugLogFile, "SrcVATSum=" & SrcVATSum & ", DstVATSum=" & DstVATSum
    DebugWriteLine DebugLogFile, "SrcTPay=" & SrcTPay & ", DstTPay=" & DstTPay
    DebugWriteLine DebugLogFile, "SrcSPay=" & SrcSPay & ", DstSPay=" & DstSPay

    If (Abs(srcQty - dstQty) > epsQty) _
       Or (Abs(SrcAmount - DstAmount) > epsMoney) _
       Or (Abs(SrcVATSum - DstVATSum) > epsMoney) _
       Or (Abs(SrcTPay - DstTPay) > epsMoney) _
       Or (Abs(SrcSPay - DstSPay) > epsMoney) Then

        DebugWriteLine DebugLogFile, "Checksum mismatch detected"
        Err.Raise vbObjectError + 779, , "Checksum reconcile failed for SessionCode=" & SessionCode
    End If

    Set rsCnt = POSCn.Execute("SELECT COUNT(*) AS Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'")
    srcCount = CLng(rsCnt.Fields("Cnt").Value)
    rsCnt.Close
    Set rsCnt = Nothing

    Set rsCnt = Cn.Execute("SELECT COUNT(*) AS Cnt FROM dbo.Transactions WHERE SessionCode='" & SessionCode & "'")
    dstCount = CLng(rsCnt.Fields("Cnt").Value)
    rsCnt.Close
    Set rsCnt = Nothing

    DebugWriteLine DebugLogFile, "srcCount=" & srcCount & ", dstCount=" & dstCount

    If dstCount <> srcCount Then
        Err.Raise vbObjectError + 777, , "Mismatch between source and server counts for SessionCode=" & SessionCode
    End If

    '========================
    ' DEBUG MODE:
    ' ·« Copied
    ' ·« „”Õ SessionCode
    ' ·« Commit ‰Â«∆Ì
    '========================
    DebugWriteLine DebugLogFile, "DEBUG MODE - data left without marking Copied"
    DebugWriteLine DebugLogFile, "DEBUG MODE - COMMIT server transaction to inspect inserted rows"
    Cn.CommitTrans

    elapsedSec = DateDiff("s", mTimeStart, Now)
    elapsedMin = elapsedSec \ 60
    elapsedSec = elapsedSec Mod 60

    frmPopup.ShowMessage "DEBUG COMPLETED" & vbCrLf & _
                         "SessionCode = " & SessionCode & vbCrLf & _
                         "Log File = " & DebugLogFile & vbCrLf & _
                         "«·Êﬁ : " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì….", vbInformation

EndSub:
    On Error Resume Next

    MousePointer = vbDefault

    If Not rsSer Is Nothing Then
        If rsSer.State = adStateOpen Then rsSer.Close
        Set rsSer = Nothing
    End If

    If Not rsPayments Is Nothing Then
        If rsPayments.State = adStateOpen Then rsPayments.Close
        Set rsPayments = Nothing
    End If

    If Not rsValueAdded Is Nothing Then
        If rsValueAdded.State = adStateOpen Then rsValueAdded.Close
        Set rsValueAdded = Nothing
    End If

    If Not rsDetails Is Nothing Then
        If rsDetails.State = adStateOpen Then rsDetails.Close
        Set rsDetails = Nothing
    End If

    If Not rsTrans Is Nothing Then
        If rsTrans.State = adStateOpen Then rsTrans.Close
        Set rsTrans = Nothing
    End If

    If Not rsCnt Is Nothing Then
        If rsCnt.State = adStateOpen Then rsCnt.Close
        Set rsCnt = Nothing
    End If

    If Not POSCn Is Nothing Then
        If POSCn.State = adStateOpen Then POSCn.Close
        Set POSCn = Nothing
    End If

    lblWait.Visible = False
    Exit Sub

ErrorHandler:
    On Error Resume Next

    DebugWriteLine DebugLogFile, "******** ERROR HANDLER ********"
    DebugWriteLine DebugLogFile, "Err.Number=" & Err.Number
    DebugWriteLine DebugLogFile, "Err.Description=" & Err.Description
    DebugWriteSQL DebugLogFile, "LastSQL At Error", LastSQL

    ' ›Ì Ê÷⁄ «·œÌ»Ã: ·« rollback
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            If Cn.Errors.Count > 0 Then
                Dim i As Long
                For i = 0 To Cn.Errors.Count - 1
                    DebugWriteLine DebugLogFile, "ADO Error #" & CStr(i + 1) & _
                        " Number=" & Cn.Errors(i).Number & _
                        ", Native=" & Cn.Errors(i).NativeError & _
                        ", Desc=" & Cn.Errors(i).Description
                Next i
            End If
        End If
    End If

    If Not POSCn Is Nothing Then
        If POSCn.State = adStateOpen Then
            If POSCn.Errors.Count > 0 Then
                Dim j As Long
                For j = 0 To POSCn.Errors.Count - 1
                    DebugWriteLine DebugLogFile, "POS ADO Error #" & CStr(j + 1) & _
                        " Number=" & POSCn.Errors(j).Number & _
                        ", Native=" & POSCn.Errors(j).NativeError & _
                        ", Desc=" & POSCn.Errors(j).Description
                Next j
            End If
        End If
    End If

    MousePointer = vbDefault
    lblWait.Visible = True
    lblWait.Caption = "DEBUG FAILED: " & Err.Description

    frmPopup.ShowMessage "DEBUG FAILED" & vbCrLf & _
                         "SessionCode = " & SessionCode & vbCrLf & _
                         "Err = " & Err.Description & vbCrLf & _
                         "Log = " & DebugLogFile, vbCritical

    GoTo EndSub

End Sub
Private Sub cmdTestUpdate_Click()

    On Error GoTo errHandler

    Dim s As String

    If ConnectionFirst = False Then Exit Sub

    s = ""
    s = s & "UPDATE DelMe "
    s = s & "SET Name = N'" & Replace("Updated " & Format$(Now, "yyyy-mm-dd hh:nn:ss"), "'", "''") & "', "
    s = s & "    Account_Code = '" & Replace("ACC-UPD", "'", "''") & "' "
    s = s & "WHERE Code = ("
    s = s & "    SELECT TOP 1 Code "
    s = s & "    FROM DelMe "
    s = s & "    WHERE Code LIKE 'T%' "
    s = s & "    ORDER BY Code DESC"
    s = s & ")"

    Cn.Execute s, , adExecuteNoRecords

    MsgBox "Update OK", vbInformation

    Exit Sub

errHandler:
    MsgBox "Update Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & s, vbCritical

End Sub
Private Sub Command2_Click()

    On Error GoTo ErrorHandler

    Dim sql As String
    Dim s As String

    Dim NoOFItem_POS As Double
    Dim NoOFItem_Server As Double
    Dim MaxItem_POS As Double
    Dim MaxItem_Server As Double

    Dim NoOfGroups_pos As Double
    Dim NoOfGroups_server As Double
    Dim MaxGroupid_pos As Double
    Dim MaxGroupidserver As Double

    Dim Rs3 As ADODB.Recordset
    Dim BatchSQL As String
    Dim BatchCount As Long

    mLastProc = "Command2_Click"
    mLastSQL = ""

    If Trim$(POSlServer.Text) = "" Then
        MsgBox "«Œ — «·‰ﬁÿ… «·„‰ﬁÊ· ≈·ÌÂ« √Ê·«", vbCritical, "OFFLINE"
        Exit Sub
    End If

    If ConnectionFirst = False Then Exit Sub

    Set Rs3 = New ADODB.Recordset

    lblWait.Visible = True

    '========================================
    ' Counts ··„—«Ã⁄… ›ﬁÿ
    '========================================
    sql = "select count(ItemID) As NoOfitems, max(ItemID) as MaxItemid from TblItems"
    Rs3.Open sql, POSConnection, adOpenStatic, adLockReadOnly, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close

    Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close

    lblWait.Caption = "Ì „ «·¬‰  ÕœÌÀ «·√”⁄«— Ê«·„·›«  «·√”«”Ì…"
    DoEvents

    sql = "select count(GroupID) As NoOfGroups, max(GroupID) as MaxGroupid from Groups"
    Rs3.Open sql, POSConnection, adOpenStatic, adLockReadOnly, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOfGroups_pos = IIf(IsNull(Rs3("NoOfGroups").Value), 0, Rs3("NoOfGroups").Value)
        MaxGroupid_pos = IIf(IsNull(Rs3("MaxGroupid").Value), 0, Rs3("MaxGroupid").Value)
    End If
    Rs3.Close

    Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOfGroups_server = IIf(IsNull(Rs3("NoOfGroups").Value), 0, Rs3("NoOfGroups").Value)
        MaxGroupidserver = IIf(IsNull(Rs3("MaxGroupid").Value), 0, Rs3("MaxGroupid").Value)
    End If
    Rs3.Close

    BolFrmLoaded = True

    '========================================
    ' 1) Groups
    '========================================
    lblWait.Caption = " ÕœÌÀ «·„Ã„Ê⁄« "
    DoEvents
    Call SyncGroups_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 2) TblUnites
    '========================================
    lblWait.Caption = " ÕœÌÀ «·ÊÕœ« "
    DoEvents
    Call SyncTblUnites_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 3) TblItems
    '========================================
    lblWait.Caption = " ÕœÌÀ «·√’‰«›"
    DoEvents
    Call SyncTblItems_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 4) TblPaymentType
    '========================================
    lblWait.Caption = " ÕœÌÀ √‰Ê«⁄ «·œ›⁄"
    DoEvents
    Call SyncTblPaymentType_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 5) TblPaymentUser
    '========================================
    lblWait.Caption = " ÕœÌÀ ’·«ÕÌ«  Ê”«∆· «·œ›⁄"
    DoEvents
    Call SyncTblPaymentUser_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 6) BanksData
    '========================================
    lblWait.Caption = " ÕœÌÀ „·› «·»‰Êﬂ"
    DoEvents
    Call SyncBanksData_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 7) TblUsers
    '========================================
    lblWait.Caption = " ÕœÌÀ «·„” Œœ„Ì‰"
    DoEvents
    Call SyncTblUsers_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 8) TblEmpJobsTypes
    '========================================
    lblWait.Caption = " ÕœÌÀ √‰Ê«⁄ «·ÊŸ«∆›"
    DoEvents
    Call SyncTblEmpJobsTypes_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 9) TblEmployee
    '========================================
    lblWait.Caption = " ÕœÌÀ «·„ÊŸ›Ì‰"
    DoEvents
    Call SyncTblEmployee_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 10) TblCustemers
    '========================================
    lblWait.Caption = " ÕœÌÀ «·⁄„·«¡"
    DoEvents
    Call SyncTblCustemers_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 11) TblStore
    '========================================
    lblWait.Caption = " ÕœÌÀ «·„Œ«“‰"
    DoEvents
    Call SyncTblStore_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 12) TblItemsUnits
    '========================================
    lblWait.Caption = " ÕœÌÀ ÊÕœ«  «·√’‰«›"
    DoEvents
    Call SyncTblItemsUnits_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 13) Update TblItemsUnits prices
    '========================================
    lblWait.Caption = " ÕœÌÀ √”⁄«— «·ÊÕœ« "
    DoEvents
    Call UpdateTblItemsUnits_ServerToPOS(BatchSQL, BatchCount)

    '========================================
    ' 14) Update TblItems
    '========================================
    lblWait.Caption = " ÕœÌÀ »Ì«‰«  «·√’‰«›"
    DoEvents
    Call UpdateTblItems_ServerToPOS(BatchSQL, BatchCount)

    lblWait.Visible = True
    lblWait.Caption = " „  ÕœÌÀ «·√”⁄«— Ê«·„·›«  «·√”«”Ì… »‰Ã«Õ"
    DoEvents

CleanExit:
    SafeCloseRS Rs3
    Exit Sub

ErrorHandler:
    LogAdoErrors POSConnection, "Command2_Click", mLastSQL, "Œÿ√ ⁄«„ ›Ì Command2_Click"
    MsgBox BuildErrMsg(POSConnection, "Command2_Click", mLastSQL, "ÕœÀ Œÿ√ √À‰«¡  ÕœÌÀ «·„·›«  «·√”«”Ì…"), vbCritical, "Œÿ√"
    lblWait.Visible = False
    SafeCloseRS Rs3
    Err.Clear

End Sub
Private Sub Command12_Click()


On Error GoTo ErrTrap

    ' «· √ﬂœ „‰ ÊÃÊœ ‰ﬁÿ… „ ’·…
    If POSlServer.Text = "" Then
        MsgBox "CIE? C????? C?????? ???C C??C", vbCritical, "OFFLINE"
        Exit Sub
    End If
    If ConnectionFirst = False Then Exit Sub

    lblWait.Visible = True

    ' «·Õ’Ê· ⁄·Ï «·ŒÌ«—«  „‰ TblOptions
    Dim rsOptions As New ADODB.Recordset
    rsOptions.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    JLCodeBasedOnBranch = IIf(rsOptions("JLCodeBasedOnBranch").Value = 0 Or IsNull(rsOptions("JLCodeBasedOnBranch").Value), False, True)
    StoreDigit = IIf(IsNull(rsOptions("StoreDigit").Value), 1, rsOptions("StoreDigit").Value)
    BranchDigit = IIf(IsNull(rsOptions("BranchDigit").Value), 1, rsOptions("BranchDigit").Value)
    rsOptions.Close

    ' ≈⁄œ«œ « ’«· «·‰ﬁÿ… (POSConnection)
    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If SysSQLServerType = 1 Then
            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer
        ElseIf SysSQLServerType = 2 Then
            If SysSQLServerTypeTechnical = "0" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & POSDb & _
                                    ";Data Source=" & POSlServer & ";Port=1433"
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                    ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
            End If
        End If
        .Open
    End With

    '  ⁄ÌÌ‰ SessionCode (’Ì€… «· «—ÌŒ Ê«·Êﬁ )
    Dim SessionCode As String
    SessionCode = Format(Now, "yyyymmddhhmmss")

    Dim mTimeStart As Date, mEndTime As Date
    mTimeStart = Now
    txtStartTime = mTimeStart
    Text3 = "Query: " & GetQuery

    ' › Õ Recordset ··„⁄«„·«  „‰ ÃœÊ· Transactions ›Ì ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ…
    Dim rsTrans As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT * FROM Transactions WHERE Copied IS NULL AND " & GetQuery
    rsTrans.CursorType = adOpenForwardOnly
    rsTrans.LockType = adLockReadOnly
    rsTrans.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    ' „ €Ì—«   Ã„Ì⁄ «·«” ⁄·«„«  (Batch)
    Dim transBatchSQL As String, detailsBatchSQL As String, valueAddedBatchSQL As String
    Dim paymentsBatchSQL As String, doubleEntryBatchSQL As String, multiPaymentBatchSQL As String, notesBatchSQL As String
    transBatchSQL = ""
    detailsBatchSQL = ""
    valueAddedBatchSQL = ""
    paymentsBatchSQL = ""
    doubleEntryBatchSQL = ""
    multiPaymentBatchSQL = ""
    notesBatchSQL = ""
    
    Dim batchThreshold As Long, recCount As Long
    batchThreshold = 50
    recCount = 0

    ' »œ¡ „⁄«„·… Ê«Õœ… ⁄·Ï ﬁ«⁄œ… «·»Ì«‰«  «·ÊÃÂ…
    Cn.BeginTrans
    Dim BeginTrans As Boolean
    BeginTrans = True

    ' «·„—Ê— ⁄·Ï ﬂ«›… «·„⁄«„·« 
    Do While Not rsTrans.EOF
         recCount = recCount + 1

         ' ﬁ—«¡… «·ﬁÌ„ „⁄ «· ⁄ÌÌ‰ «·«› —«÷Ì ﬂ„« ›Ì «·ﬂÊœ «·√’·Ì
         Dim PayMentType As Long, cusID As Long, BranchID As Long, BoxID As Long, BillBasedOn As Double
         Dim VAT As Double, VATYou As Double, NoteId As Long, Trans_DiscountType As Long
         Dim Trans_Discount As Double, TaxValue As Double, order_no As String, SaleType As Long
         Dim TaxAddValue As Double, NetValue As Double, Transaction_NetValue As Double, DepandToConv As Long
         Dim CarTypeID As Long, OilsTypesID As Long, YearFact As Long, FixesAssetsID As Long, ColorID2 As Long
         Dim KM As Double, PPointID As Long, SupplerID As Long, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double
         Dim CarEnginoil As Double, CarGearOil As Double, InvoiceTypeCodeID As Long
         Dim storeID As Variant, userID As Variant, Emp_ID As Variant
         Dim NoteSerial As String, NoteSerial1 As String, TransactionComment As String
         Dim CashCustomerName As String, CashCustomerPhone As String
         Dim PlateNo As String, Shaseh As String, CarMeter As String
         Dim CIBAN As String, RecTime As Date, ActualDeliveryDate As Date, LatestDeliveryDate As Date
         Dim InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String
         Dim paymentnote As String, PaymentMeansCode As String
         Dim FromTransaction_Date As Date
         
         
Dim CarOilChangeDate As Date
         PayMentType = Val(rsTrans("PaymentType").Value & "")
         cusID = Val(rsTrans("CusID").Value & "")
         storeID = rsTrans("StoreID").Value
         userID = rsTrans("UserID").Value
         Emp_ID = rsTrans("Emp_ID").Value
         BranchID = Val(rsTrans("BranchID").Value & "")
         BoxID = Val(rsTrans("BoxID").Value & "")
         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
         VAT = Val(rsTrans("VAT").Value & "")
         VATYou = Val(rsTrans("VATYou").Value & "")
         NoteSerial = rsTrans("NoteSerial").Value
         NoteSerial1 = rsTrans("NoteSerial1").Value
         NoteId = Val(rsTrans("NoteId").Value & "")
         FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
         TransactionComment = rsTrans("TransactionComment").Value & ""
         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
         FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
         TaxValue = Val(rsTrans("TaxValue").Value & "")
         order_no = Val(rsTrans("order_no").Value & "")
         SaleType = Val(rsTrans("SaleType").Value & "")
         CashCustomerName = rsTrans("CashCustomerName").Value & ""
         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
         NetValue = Val(rsTrans("NetValue").Value & "")
         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
         PlateNo = rsTrans("PlateNo").Value & ""
         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
         YearFact = Val(rsTrans("YearFact").Value & "")
         Shaseh = rsTrans("Shaseh").Value & ""
         CarMeter = rsTrans("CarMeter").Value & ""
         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
         ColorID2 = Val(rsTrans("ColorID2").Value & "")
         KM = Val(rsTrans("KM").Value & "")
         Chasee = rsTrans("Chasee").Value & ""
         PPointID = Val(rsTrans("PPointID").Value & "")
         Phone2 = rsTrans("Phone2").Value & ""
         SupplerID = Val(rsTrans("SupplerID").Value & "")
         Ser = Val(rsTrans("Ser").Value & "")
         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
         ' ≈–« ·„ Ì „  ÕœÌœ ﬁÌ„… ·‹ CarOilChangeDate° ‰” Œœ„ «· «—ÌŒ «·Õ«·Ì
         If Trim(rsTrans("CarOilChangeDate").Value & "") = "" Then
             CarOilChangeDate = Date
         Else
             CarOilChangeDate = rsTrans("CarOilChangeDate").Value & ""
         End If
         CIBAN = rsTrans("CIBAN").Value & ""
         RecTime = rsTrans("RecTime").Value & ""
         ActualDeliveryDate = rsTrans("ActualDeliveryDate").Value & ""
         LatestDeliveryDate = rsTrans("LatestDeliveryDate").Value & ""
         InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
         InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
         DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
         TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
         paymentnote = rsTrans("paymentnote").Value & ""
         PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""

         ' ﬁ—«¡…  «—ÌŒ «·„⁄«„·… „‰ «·”Ã· («› —÷ √‰Â „ÊÃÊœ)
         FromTransaction_Date = rsTrans("Transaction_Date").Value

         ' ≈–« ﬂ«‰ POSBillType = 0°  ⁄œÌ· NoteSerial ÊNoteId
         If Val(rsTrans("POSBillType").Value & "") = 0 Then
             NoteSerial = Notes_coding(CInt(BranchID), FromTransaction_Date)
             NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
         End If

         '  ⁄œÌ· TransactionComment ﬂ„« ›Ì «·ﬂÊœ «·√’·Ì
         TransactionComment = " ?CE??E ?????E ?? ?C?IE  " & POSname.Text & "   " & _
                              "  ??? C??CE??E  C?C???E" & NoteSerial1

         '  Ê·Ìœ —ﬁ„ ÃœÌœ ··„⁄«„·… ⁄·Ï «·ÊÃÂ…
         Dim currentDestTransactionID As String
         currentDestTransactionID = CStr(new_id("Transactions", "Transaction_ID", "", True))

         ' »‰«¡ «” ⁄·«„ «·≈œ—«Ã «·ﬂ«„· ··„⁄«„·… „⁄ ﬂ«›… «·ÕﬁÊ·
         Dim transSQL As String
         transSQL = "INSERT INTO [" & ServerDb & "].dbo.Transactions (" & _
                    "Transaction_ID, Transaction_Date, TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
                    "CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, " & _
                    "NoteSerial1, NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, " & _
                    "OldNoteserial, OldNoteId, OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, " & _
                    "SaleType, CashCustomerName, TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, " & _
                    "DepandToConv, CarTypeID, PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, " & _
                    "ColorID2, KM, Chasee, PPointID, Phone2, SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, " & _
                    "CarGearOil, CarOilChangeDate, CIBAN, RecTime, ActualDeliveryDate, LatestDeliveryDate, " & _
                    "InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("
         transSQL = transSQL & currentDestTransactionID & "," & SQLDate(FromTransaction_Date, True) & "," & _
                    Val(rsTrans("TypeInvoice").Value & "") & ",'" & rsTrans("Transaction_Serial").Value & "'," & _
                    FromTransaction_Type & "," & PayMentType & "," & Val(cusID) & "," & storeID & "," & userID & "," & _
                    Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & "," & _
                    "'" & NoteSerial & "'," & _
                    "'" & NoteSerial1 & "',0,1," & _
                    "'" & TransactionComment & "','" & SessionCode & "'," & Val(rsTrans("POSBillType").Value & "") & "," & _
                    "'" & rsTrans("Noteserial1").Value & "','" & Trim(rsTrans("Noteserial").Value & "") & "'," & _
                    Val(rsTrans("NoteId").Value & "") & "," & rsTrans("Transaction_ID").Value & "," & _
                    Trans_DiscountType & "," & Trans_Discount & "," & TaxValue & ",'" & order_no & "'," & _
                    SaleType & ",'" & CashCustomerName & "'," & TaxAddValue & ",'" & CashCustomerPhone & "'," & _
                    SQLDate(rsTrans("last_changed").Value, True) & "," & NetValue & "," & Transaction_NetValue & "," & IIf(DepandToConv, 1, 0) & "," & _
                    CarTypeID & ",'" & PlateNo & "'," & OilsTypesID & "," & YearFact & ",'" & Shaseh & "','" & CarMeter & "'," & _
                    FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Chasee & "'," & PPointID & ",'" & Phone2 & "'," & _
                    SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & CarEnginoil & "," & _
                    CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & CIBAN & "'," & SQLDate(RecTime, True) & "," & _
                    SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
                    InvoiceTypeCodeID & ",'" & InvoiceTypeCodename & "','" & DocumentCurrencyCode & "'," & _
                    "'" & TaxCurrencyCode & "','" & paymentnote & "','" & PaymentMeansCode & "')"
         
         '  Ã„Ì⁄ «” ⁄·«„ «·„⁄«„·…
         transBatchSQL = transBatchSQL & transSQL & vbCrLf

         ' --- „⁄«·Ã…  ›«’Ì· «·„⁄«„·… (Transaction_Details) ---
         Dim rsDetails As New ADODB.Recordset
         sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & FromTransaction_ID
         rsDetails.CursorType = adOpenForwardOnly
         rsDetails.LockType = adLockReadOnly
         rsDetails.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
         Do While Not rsDetails.EOF
              Dim detailSQL As String
              detailSQL = "INSERT INTO [" & ServerDb & "].dbo.Transaction_Details (" & _
                          "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
                          "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
                          "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES ("
              detailSQL = detailSQL & currentDestTransactionID & "," & rsDetails("Item_ID").Value & "," & rsDetails("ItemCase").Value & "," & _
                          rsDetails("Quantity").Value & "," & rsDetails("Price").Value & "," & rsDetails("ItemDiscountType").Value & "," & _
                          rsDetails("ItemDiscount").Value & "," & rsDetails("ShowQty").Value & "," & rsDetails("showPrice").Value & "," & _
                          rsDetails("UnitId").Value & "," & rsDetails("ColorID").Value & "," & rsDetails("ItemSize").Value & "," & _
                          rsDetails("ClassId").Value & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo").Value & "") & "," & Val(rsDetails("PumpId").Value & "") & "," & _
                          Val(rsDetails("PrevQty").Value & "") & ",'" & rsDetails("PrintName").Value & "'," & Val(rsDetails("Cash").Value & "") & "," & _
                          Val(rsDetails("Mada").Value & "") & "," & Val(rsDetails("Visa").Value & "") & "," & Val(rsDetails("Deferred").Value & "") & "," & _
                          Val(rsDetails("AmountH").Value & "") & "," & Val(rsDetails("AmountHComm").Value & "") & ",'" & rsDetails("DetailsPump").Value & "'," & _
                          "'" & rsDetails("Account_CodeComm").Value & "','" & rsDetails("Account_Code").Value & "'," & _
                          IIf(rsDetails("IsOther").Value, 1, 0) & ")"
              detailsBatchSQL = detailsBatchSQL & detailSQL & vbCrLf
              rsDetails.MoveNext
         Loop
         rsDetails.Close

         ' --- „⁄«·Ã… TransactionValueAdded ---
         Dim rsValueAdded As New ADODB.Recordset
         sql = "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & FromTransaction_ID
         rsValueAdded.CursorType = adOpenForwardOnly
         rsValueAdded.LockType = adLockReadOnly
         rsValueAdded.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
         Do While Not rsValueAdded.EOF
              Dim valueSQL As String
              valueSQL = "INSERT INTO [" & ServerDb & "].dbo.TransactionValueAdded (" & _
                         "Transaction_ID, ItemID, Vatyo, VAT, Valu, selectd, Transaction_Type, SessionCode) VALUES ("
              valueSQL = valueSQL & currentDestTransactionID & "," & rsValueAdded("ItemID").Value & "," & rsValueAdded("Vatyo").Value & "," & _
                         rsValueAdded("Vat").Value & "," & rsValueAdded("Valu").Value & "," & rsValueAdded("selectd").Value & "," & _
                         rsValueAdded("Transaction_Type").Value & ",'" & SessionCode & "')"
              valueAddedBatchSQL = valueAddedBatchSQL & valueSQL & vbCrLf
              rsValueAdded.MoveNext
         Loop
         rsValueAdded.Close

         ' --- „⁄«·Ã… TblTransactionPayments ≈–« ﬂ«‰ ‰Ê⁄ «·„⁄«„·… 21 √Ê 9 ---
         If FromTransaction_Type = 21 Or FromTransaction_Type = 9 Then
             Dim rsPayments As New ADODB.Recordset
             sql = "SELECT * FROM TblTransactionPayments WHERE Transaction_ID = " & FromTransaction_ID
             rsPayments.CursorType = adOpenForwardOnly
             rsPayments.LockType = adLockReadOnly
             rsPayments.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
             Do While Not rsPayments.EOF
                  Dim paymentSQL As String, Recorddate As Date
                  Recorddate = IIf(IsNull(rsPayments("Recorddate").Value), Now, rsPayments("Recorddate").Value)
                  paymentSQL = "INSERT INTO [" & ServerDb & "].dbo.TblTransactionPayments (" & _
                               "Transaction_ID, boxid, Recorddate, PointID, CurrentCashireID, PaymentID, Value, CardNo, Effect, SessionCode) VALUES ("
                  paymentSQL = paymentSQL & currentDestTransactionID & "," & rsPayments("boxid").Value & "," & SQLDate(Recorddate, True) & "," & _
                               rsPayments("PointID").Value & "," & rsPayments("CurrentCashireID").Value & "," & rsPayments("PaymentID").Value & "," & _
                               rsPayments("Value").Value & ",'" & rsPayments("CardNo").Value & "'," & rsPayments("Effect").Value & ",'" & SessionCode & "')"
                  paymentsBatchSQL = paymentsBatchSQL & paymentSQL & vbCrLf
                  rsPayments.MoveNext
             Loop
             rsPayments.Close
         End If

         ' --- ≈–« POSBillType = 0 And Transaction_Type <> 42° „⁄«·Ã… Notes, DOUBLE_ENTREY_VOUCHERS Ê TblMultuPayment ---
         If Val(rsTrans("POSBillType").Value & "") = 0 And FromTransaction_Type <> 42 Then
'             ' „⁄«·Ã… Notes
'             Dim noteSQL As String
'             noteSQL = "INSERT INTO [" & ServerDb & "].dbo.Notes (" & _
'                       "NoteID, NoteDate, NoteType, NoteSerial, NoteSerial1, branch_no, Transaction_ID, UserID, SessionCode) VALUES ("
'             noteSQL = noteSQL & NoteId & "," & SQLDate(FromTransaction_Date, True) & "," & mNoteType & ",'" & NoteSerial & "'," & _
'                       "'" & NoteSerial1 & "'," & BranchID & "," & currentDestTransactionID & ",1,'" & SessionCode & "')"
'             notesBatchSQL = notesBatchSQL & noteSQL & vbCrLf
'
'             ' „⁄«·Ã… DOUBLE_ENTREY_VOUCHERS
'             Dim rsDoubleEntry As New ADODB.Recordset
'             sql = "SELECT * FROM DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID = " & rsTrans("NoteId").Value
'             rsDoubleEntry.CursorType = adOpenForwardOnly
'             rsDoubleEntry.LockType = adLockReadOnly
'             rsDoubleEntry.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'             Do While Not rsDoubleEntry.EOF
'                  Dim doubleSQL As String, DEVID As String
'                  DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
'                  doubleSQL = "INSERT INTO [" & ServerDb & "].dbo.DOUBLE_ENTREY_VOUCHERS (" & _
'                              "Double_Entry_Vouchers_ID, DEV_ID_Line_No, Account_Code, Value, Credit_Or_Debit, " & _
'                              "Double_Entry_Vouchers_Description, RecordDate, Notes_ID, branch_id, UserID, Transaction_ID, SessionCode) VALUES ("
'                  doubleSQL = doubleSQL & DEVID & "," & rsDoubleEntry("DEV_ID_Line_No").Value & ",'" & rsDoubleEntry("Account_Code").Value & "'," & _
'                              rsDoubleEntry("Value").Value & "," & rsDoubleEntry("Credit_Or_Debit").Value & ",'" & _
'                              rsDoubleEntry("Double_Entry_Vouchers_Description").Value & "'," & SQLDate(FromTransaction_Date, True) & "," & _
'                              rsTrans("NoteId").Value & "," & rsDoubleEntry("branch_id").Value & ",1," & currentDestTransactionID & ",'" & SessionCode & "')"
'                  doubleEntryBatchSQL = doubleEntryBatchSQL & doubleSQL & vbCrLf
'                  rsDoubleEntry.MoveNext
'             Loop
'             rsDoubleEntry.Close

             ' „⁄«·Ã… TblMultuPayment
             Dim rsMultiPay As New ADODB.Recordset
             sql = "SELECT * FROM TblMultuPayment WHERE NoteID = " & rsTrans("NoteId").Value
             rsMultiPay.CursorType = adOpenForwardOnly
             rsMultiPay.LockType = adLockReadOnly
             rsMultiPay.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
             Do While Not rsMultiPay.EOF
                  Dim multiSQL As String
                  multiSQL = "INSERT INTO [" & ServerDb & "].dbo.TblMultuPayment (" & _
                             "NoteId, PaymentID, Value, CardNo, maxvalue, SessionCode) VALUES ("
                  multiSQL = multiSQL & rsMultiPay("NoteId").Value & "," & rsMultiPay("PaymentID").Value & "," & rsMultiPay("Value").Value & ",'" & _
                             rsMultiPay("CardNo").Value & "'," & rsMultiPay("maxvalue").Value & ",'" & SessionCode & "')"
                  multiPaymentBatchSQL = multiPaymentBatchSQL & multiSQL & vbCrLf
                  rsMultiPay.MoveNext
             Loop
             rsMultiPay.Close
         End If

         ' «·«‰ ﬁ«· ··”Ã· «· «·Ì
         rsTrans.MoveNext

         '  ‰›Ì– «·œı›⁄… ≈–« Ê’· ⁄œœ «·”Ã·«  ··Õœ «·„Õœœ
         If recCount Mod batchThreshold = 0 Then
             If transBatchSQL <> "" Then Cn.Execute transBatchSQL: transBatchSQL = ""
             If detailsBatchSQL <> "" Then Cn.Execute detailsBatchSQL: detailsBatchSQL = ""
             If valueAddedBatchSQL <> "" Then Cn.Execute valueAddedBatchSQL: valueAddedBatchSQL = ""
             If paymentsBatchSQL <> "" Then Cn.Execute paymentsBatchSQL: paymentsBatchSQL = ""
        '     If doubleEntryBatchSQL <> "" Then Cn.Execute doubleEntryBatchSQL: doubleEntryBatchSQL = ""
             If multiPaymentBatchSQL <> "" Then Cn.Execute multiPaymentBatchSQL: multiPaymentBatchSQL = ""
        '     If notesBatchSQL <> "" Then Cn.Execute notesBatchSQL: notesBatchSQL = ""
         End If

    Loop
    rsTrans.Close

    '  ‰›Ì– √Ì œ›⁄«  „ »ﬁÌ…
    If transBatchSQL <> "" Then Cn.Execute transBatchSQL
    If detailsBatchSQL <> "" Then Cn.Execute detailsBatchSQL
    If valueAddedBatchSQL <> "" Then Cn.Execute valueAddedBatchSQL
    If paymentsBatchSQL <> "" Then Cn.Execute paymentsBatchSQL
   ' If doubleEntryBatchSQL <> "" Then Cn.Execute doubleEntryBatchSQL
    If multiPaymentBatchSQL <> "" Then Cn.Execute multiPaymentBatchSQL
  '  If notesBatchSQL <> "" Then Cn.Execute notesBatchSQL

    '  ÕœÌÀ ”Ã·«  Transactions ›Ì ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ… ( ⁄ÌÌ‰ Copied = 1 „⁄ SessionCode)
    sql = "UPDATE [" & POSDb & "].dbo.Transactions SET Copied = 1, SessionCode = '" & SessionCode & "' WHERE Copied IS NULL AND " & GetQuery
    POSConnection.Execute sql

    ' ≈–« ﬂ«‰ chkRec „›⁄·°  ÕœÌÀ ÃœÊ· Notes œ›⁄… Ê«Õœ…
'    If chkRec.Value = vbChecked Then
'         sql = "UPDATE [" & POSDb & "].dbo.Notes SET Copied = 1, SessionCode = '" & SessionCode & "' WHERE NoteType = 4 AND Copied IS NULL AND NoteDate = " & SQLDate(dbRecordDate.Value, False)
'         POSConnection.Execute sql
'    End If

    '  ”ÃÌ· ”Ã· ›Ì TblOffline · ÊÀÌﬁ ⁄„·Ì… «·‰ﬁ·
    Dim rsOffline As New ADODB.Recordset
    sql = "SELECT * FROM TblOffline WHERE 1 = -1"
    rsOffline.Open sql, Cn, adOpenKeyset, adLockOptimistic
    rsOffline.AddNew
    rsOffline!Recorddate = Date
    rsOffline!StartTime = mTimeStart
    mEndTime = Now
    rsOffline!EndTime = mEndTime
    rsOffline!SessionCode = SessionCode
    rsOffline!POSname = POSlServer.Text
    rsOffline!CountSalesOfeers = CountSalesOfeers
    rsOffline!CountSales = CountSales
    rsOffline!CountSalesReturn = CountSalesReturn
    rsOffline!CountPurchase = CountPurchase
    rsOffline!CountPurchaseReturn = CountPurchaseReturn
    rsOffline!CountRec = CountRec
    rsOffline.Update
    rsOffline.Close

    ' ≈‰Â«¡ «·„⁄«„·…
    Cn.CommitTrans
    BeginTrans = False

    lblWait.Visible = False
    txtEndTime = mEndTime
    txtCountSalesReturn = CountSalesReturn
    txtCountSales = CountSales
    txtCountSalesOfeers = CountSalesOfeers
    Exit Sub

ErrTrap:
    If BeginTrans Then
         Cn.RollbackTrans
         BeginTrans = False
    End If
    If Err.Number = -2147217900 Then
         MsgBox "?C ???? ??U ??? C?E?C?CE" & vbCrLf & "??I E? CIIC? ??? U?? ?C??E", vbExclamation, App.Title
         Exit Sub
    End If
    MsgBox "Œÿ√: " & Err.Description, vbExclamation, App.Title


End Sub


'Private Sub Command14_Click()
'
'
'    On Error GoTo ErrorHandler
'
'    ' «· √ﬂœ „‰ ÊÃÊœ ‰ﬁÿ… „ ’·…
'    If POSlServer.Text = "" Then
'        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
'        Exit Sub
'    End If
'    If ConnectionFirst = False Then Exit Sub
'
'
'  lblWait.Visible = True
'lblWait.Caption = "Ã«—Ì »œ¡ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄« ..."
'DoEvents
'MousePointer = vbHourglass
'WritePhaseLog "Start Command14_Click"
'
'Dim rsCnt As ADODB.Recordset
'
'    Dim POSConnection As New ADODB.Connection
'
'    Dim rsTrans As New ADODB.Recordset
'    Dim rsDetails As New ADODB.Recordset
'    Dim rsValueAdded As New ADODB.Recordset
'    Dim rsPayments As New ADODB.Recordset
'
'    Dim BatchSize As Integer, recCounter As Integer
'    BatchSize = 50
'    recCounter = 0
'Dim batchThreshold As Long, recCount As Long
'    batchThreshold = 50
'    recCount = 0
'
'    Dim transBatchSQL As String, detailsBatchSQL As String, valueAddedBatchSQL As String, paymentsBatchSQL2 As String, paymentsBatchSQL As String, LastSQL As String
'    transBatchSQL = ""
'    detailsBatchSQL = ""
'    valueAddedBatchSQL = ""
'
'    Dim SessionCode As String
'    SessionCode = Format(Now, "yyyymmddhhmmss")
'UpdateTransferCaption "Ã«—Ì  ÃÂÌ“ Ã·”… «·‰ﬁ·", 0, 0, SessionCode, Now
'    ' › Õ « ’«· POS
'    POSConnection.CursorLocation = adUseServer
'    POSConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
'    POSConnection.Open
'
'    ' › Õ « ’«· «·”Ì—›— «·„—ﬂ“Ì
'    Cn.CursorLocation = adUseServer
''    Cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & mDBPOSName & ";Data Source=RemoteServer10"
''    Cn.Open
'
'    ' › Õ Recordset ··„⁄«„·«  „‰ POS
'
'
'    POSConnection.Execute "UPDATE Transactions SET SessionCode = '" & SessionCode & "' WHERE IsNull(Copied,0) =0 AND " & GetQuery
'WritePhaseLog "Tagging source transactions", "SessionCode=" & SessionCode
'UpdateTransferCaption " „  ⁄·Ì„ «·›Ê« Ì— «·„—«œ ‰ﬁ·Â«", 0, 0, SessionCode, mTimeStart
'
'    rsTrans.Open "SELECT * FROM Transactions WHERE IsNull(Copied,0) =0 AND SessionCode = '" & SessionCode & "' AND " & GetQuery & " ORDER BY Transaction_ID", POSConnection, adOpenForwardOnly, adLockReadOnly
'
'
'Dim TotalInvoices As Long
'Dim CurrentInvoiceNo As String
'Dim BatchNo As Long
'BatchNo = 0
'
'Set rsCnt = POSConnection.Execute( _
'    "SELECT COUNT(*) AS Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "' AND " & GetQuery)
'TotalInvoices = CLng(rsCnt!Cnt)
'rsCnt.Close
'Set rsCnt = Nothing
'
'UpdateTransferCaption " „ «·⁄ÀÊ— ⁄·Ï ›Ê« Ì— ··‰ﬁ·", 0, TotalInvoices, SessionCode, mTimeStart
'WritePhaseLog "Invoices selected", "Count=" & TotalInvoices
'
'    If rsTrans.EOF Then
'        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
'        GoTo EndSub
'    End If
'
'
'Dim CarOilChangeDate As Date, RecTime As Date, mTimeIn As String
'
'    '
'    mTimeStart = Now
'    txtStartTime = mTimeStart
'    Text3 = "Query: " & GetQuery
'    Cn.Execute "SET XACT_ABORT ON;"
'    Cn.BeginTrans
'
'    Do While Not rsTrans.EOF
'
'        If Trim(rsTrans("NoteSerial1") & "") <> "" Then
'            CurrentInvoiceNo = Trim(rsTrans("NoteSerial1") & "")
'        Else
'            CurrentInvoiceNo = CStr(Val(rsTrans("Transaction_ID") & ""))
'        End If
'
'        If recCounter = 0 Or recCounter Mod 5 = 0 Then
'            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·›« Ê—… —ﬁ„ " & CurrentInvoiceNo, recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'        End If
'
'
''        Dim currentDestTransactionID As Long
''        currentDestTransactionID = new_id("Transactions", "Transaction_ID", "", True) + recCounter
''
''        '  Ã„Ì⁄ Transaction
''        transBatchSQL = transBatchSQL & "INSERT INTO Transactions (Transaction_ID, Transaction_Date, Transaction_Type, PaymentType, CusID, StoreID, Emp_ID, BranchID, SessionCode) VALUES (" & _
''            currentDestTransactionID & "," & SQLDate(rsTrans("Transaction_Date"), True) & "," & rsTrans("Transaction_Type") & "," & rsTrans("PaymentType") & "," & rsTrans("CusID") & "," & rsTrans("StoreID") & "," & rsTrans("Emp_ID") & "," & rsTrans("BranchID") & ",'" & SessionCode & "');"
'    recCount = recCount + 1
'   ' ﬁ—«¡… «·ÕﬁÊ· „⁄ «·ﬁÌ„ «·«› —«÷Ì… ﬂ„« ›Ì «·ﬂÊœ «·√’·Ì
'         Dim PayMentType As Long, CusID As Long, BranchID As Integer, BoxID As Long, BillBasedOn As Double
'         Dim VAT As Double, VATYou As Double, NoteId As Long, Trans_DiscountType As Long
'         Dim Trans_Discount As Double, TaxValue As Double, order_no As Long, SaleType As Long
'         Dim TaxAddValue As Double, NetValue As Double, Transaction_NetValue As Double, DepandToConv As Long
'         Dim CarTypeID As Long, OilsTypesID As Long, YearFact As Long, FixesAssetsID As Long, ColorID2 As Long
'         Dim KM As Double, PPointID As Long, SupplerID As Long, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double
'         Dim CarEnginoil As Double, CarGearOil As Double, InvoiceTypeCodeID As Long
'         Dim StoreId As Variant, UserID As Variant, Emp_ID As Variant
'         Dim NoteSerial As String, NoteSerial1 As String, TransactionComment As String
'         Dim CashCustomerName As String, CashCustomerPhone As String
'         Dim PlateNo As String, Shaseh As String, CarMeter As String
'         Dim CIBAN As String
'         Dim InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String
'         Dim paymentnote As String, PaymentMeansCode As String
'         Dim FromTransaction_Date As Date
'
'         PayMentType = Val(rsTrans("PaymentType").Value & "")
'          FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
'         CusID = Val(rsTrans("CusID").Value & "")
'         StoreId = Val(rsTrans("StoreID").Value & "")
'         UserID = Val(rsTrans("UserID").Value & "")
'         Emp_ID = Val(rsTrans("Emp_ID").Value & "")
'         BranchID = Val(rsTrans("BranchID").Value & "")
'         BoxID = Val(rsTrans("BoxID").Value & "")
'         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
'          PayMentType = Val(rsTrans("PaymentType").Value & "")
'         CusID = Val(rsTrans("CusID").Value & "")
'         StoreId = Val(rsTrans("StoreID").Value & "")
'         UserID = Val(rsTrans("UserID").Value & "")
'         Emp_ID = Val(rsTrans("Emp_ID").Value & "")
'         BranchID = Val(rsTrans("BranchID").Value & "")
'         BoxID = Val(rsTrans("BoxID").Value & "")
'         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
'         mTimeIn = Trim(rsTrans("TimeIn").Value & "")
'         VAT = Val(rsTrans("VAT").Value & "")
'         VATYou = Val(rsTrans("VATYou").Value & "")
'         NoteSerial = rsTrans("NoteSerial").Value & ""
'         NoteSerial1 = rsTrans("NoteSerial1").Value & ""
'         NoteId = Val(rsTrans("NoteId").Value & "")
'         FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
'         TransactionComment = rsTrans("TransactionComment").Value & ""
'         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
'         FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
'         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
'         TaxValue = Val(rsTrans("TaxValue").Value & "")
'         order_no = Val(rsTrans("order_no").Value & "")
'         SaleType = Val(rsTrans("SaleType").Value & "")
'         CashCustomerName = rsTrans("CashCustomerName").Value & ""
'         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
'         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
'         NetValue = Val(rsTrans("NetValue").Value & "")
'         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
'         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
'         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
'         PlateNo = rsTrans("PlateNo").Value & ""
'         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
'         YearFact = Val(rsTrans("YearFact").Value & "")
'         Shaseh = rsTrans("Shaseh").Value & ""
'         CarMeter = rsTrans("CarMeter").Value & ""
'         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
'         ColorID2 = Val(rsTrans("ColorID2").Value & "")
'         KM = Val(rsTrans("KM").Value & "")
'         Chasee = rsTrans("Chasee").Value & ""
'         PPointID = Val(rsTrans("PPointID").Value & "")
'         Phone2 = rsTrans("Phone2").Value & ""
'         SupplerID = Val(rsTrans("SupplerID").Value & "")
'         Ser = Val(rsTrans("Ser").Value & "")
'         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
'         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
'         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
'         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
'         VAT = Val(rsTrans("VAT").Value & "")
'         VATYou = Val(rsTrans("VATYou").Value & "")
'         NoteSerial = rsTrans("NoteSerial").Value & ""
'         NoteSerial1 = rsTrans("NoteSerial1").Value & ""
'         NoteId = Val(rsTrans("NoteId").Value & "")
'         TransactionComment = rsTrans("TransactionComment").Value & ""
'         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
'         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
'         TaxValue = Val(rsTrans("TaxValue").Value & "")
'         order_no = Val(rsTrans("order_no").Value & "")
'         SaleType = Val(rsTrans("SaleType").Value & "")
'         CashCustomerName = rsTrans("CashCustomerName").Value & ""
'         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
'         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
'         NetValue = Val(rsTrans("NetValue").Value & "")
'         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
'         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
'         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
'         PlateNo = rsTrans("PlateNo").Value & ""
'         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
'         YearFact = Val(rsTrans("YearFact").Value & "")
'         Shaseh = rsTrans("Shaseh").Value & ""
'         CarMeter = rsTrans("CarMeter").Value & ""
'         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
'         ColorID2 = Val(rsTrans("ColorID2").Value & "")
'         KM = Val(rsTrans("KM").Value & "")
'         Chasee = rsTrans("Chasee").Value & ""
'         PPointID = Val(rsTrans("PPointID").Value & "")
'         Phone2 = rsTrans("Phone2").Value & ""
'         SupplerID = Val(rsTrans("SupplerID").Value & "")
'         Ser = Val(rsTrans("Ser").Value & "")
'         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
'         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
'         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
'         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
'         If Trim(rsTrans("CarOilChangeDate").Value & "") = "" Then
'             CarOilChangeDate = Date
'         Else
'             CarOilChangeDate = rsTrans("CarOilChangeDate").Value & ""
'         End If
'         CIBAN = rsTrans("CIBAN").Value & ""
'         'RecTime = IIf(rsTrans("RecTime").Value & "" = "", Time, rsTrans("RecTime").Value & "")
'         Dim tmpRecTime As Variant
'tmpRecTime = rsTrans("RecTime").Value
''If IsNull(tmpRecTime) Or Trim(CStr(tmpRecTime)) = "" Or tmpRecTime = "#12/30/1899#" Then
''    RecTime = Time
''Else
''    RecTime = tmpRecTime
''End If
''RecTime = IIf(IsNull(rsTrans("RecTime").Value), Now, rsTrans("RecTime").Value)
'
'Dim tmpRecTimeStr As String
''tmpRecTimeStr = Trim(CStr(rsTrans("RecTime").Value & ""))
''
''If tmpRecTimeStr = "" Or tmpRecTimeStr = "30-Dec-1899" Then
''    RecTime = Time
''ElseIf IsDate(tmpRecTimeStr) Then
''    RecTime = CDate(tmpRecTimeStr)
''Else
''    RecTime = Time
''End If
'
'Dim v As Variant: v = rsTrans("RecTime").Value
'If IsDate(v) Then
'    If Year(CDate(v)) = 1899 And Month(CDate(v)) = 12 And Day(CDate(v)) = 30 Then
'        RecTime = Time
'    Else
'        RecTime = CDate(v)
'    End If
'Else
'    RecTime = Time
'End If
'
'
'        ' RecTime = IIf(IsNull(rsTrans("RecTime").Value) Or Trim(rsTrans("RecTime").Value & "") = "", Time, rsTrans("RecTime").Value)
'         ActualDeliveryDate = IIf(rsTrans("ActualDeliveryDate").Value & "" = "", Date, rsTrans("ActualDeliveryDate").Value & "")
'         LatestDeliveryDate = IIf(rsTrans("LatestDeliveryDate").Value & "" = "", Date, rsTrans("ActualDeliveryDate").Value & "")
'
'         InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
'         InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
'         DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
'         TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
'         paymentnote = rsTrans("paymentnote").Value & ""
'         PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""
'         FromTransaction_Date = rsTrans("Transaction_Date").Value
'
'         ' ≈–« POSBillType = 0°  ⁄œÌ· NoteSerial ÊNoteId
'         If Val(rsTrans("POSBillType").Value & "") = 0 Then
'             NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
'             NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
'         End If
'
'         TransactionComment = " ›« Ê—… „‰ﬁÊ·… „‰ " & POSname.Text & "   " & _
'                              "   —ﬁ„ «·›« Ê—… " & NoteSerial1
'   '  Ê·Ìœ —ﬁ„ ÃœÌœ ··„⁄«„·… ⁄·Ï «·ÊÃÂ…
''         Dim currentDestTransactionID As String
''         currentDestTransactionID = CStr((new_id("Transactions", "Transaction_ID", "", True) + recCount))
'        Dim currentDestTransactionID As String
'        Dim rsSer As ADODB.Recordset
'        LastSQL = "EXEC dbo.ReserveTransactionId"
'If recCounter = 0 Or recCounter Mod 10 = 0 Then
'    UpdateTransferCaption "Ã«—Ì ÕÃ“ √—ﬁ«„ «·„⁄«„·«  ⁄·Ï «·”Ì—›—", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'End If
'        ' ‰œ«¡ «·” Ê—œ »—Ê”ÌÃ—
'        Set rsSer = Cn.Execute("EXEC dbo.ReserveTransactionId")
'
'
'        If Not (rsSer.EOF) Then
'            currentDestTransactionID = CStr(rsSer.Fields("NewId").Value)
'        Else
'            Err.Raise vbObjectError + 500, , "·„ Ì „ ≈—Ã«⁄ Transaction_ID ÃœÌœ „‰ «·”Ì—›—"
'        End If
'
'        rsSer.Close
'        Set rsSer = Nothing
'
'
'transSQL = "INSERT INTO " & mServerD & "Transactions (" & _
'"Transaction_ID, Transaction_Date,TimeIn ,TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
'"CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, NoteSerial1, " & _
'"NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, OldNoteserial, OldNoteId, " & _
'"OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, SaleType, CashCustomerName, " & _
'"TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, DepandToConv, CarTypeID, " & _
'"PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, ColorID2, KM, Chasee, PPointID, Phone2, " & _
'"SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, CarGearOil, CarOilChangeDate, CIBAN, RecTime, " & _
'"ActualDeliveryDate, LatestDeliveryDate, InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, " & _
'"TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("
'
'transSQL = transSQL & currentDestTransactionID & "," & SQLDate(FromTransaction_Date, True) & ",'" & Trim(mTimeIn) & "'," & _
'Val(rsTrans("TypeInvoice") & "") & ",'" & Replace(rsTrans("Transaction_Serial") & "", "'", "''") & "'," & _
'FromTransaction_Type & "," & PayMentType & "," & CusID & "," & StoreId & "," & UserID & "," & _
'Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & ",'" & _
'NoteSerial & "','" & NoteSerial1 & "'," & NoteId & ",1,'" & Replace(TransactionComment, "'", "''") & "','" & _
'SessionCode & "'," & IIf(Val(rsTrans("POSBillType") & "") = 0, 1, Val(rsTrans("POSBillType") & "")) & ",'" & rsTrans("Noteserial1") & "" & "','" & Trim(rsTrans("Noteserial") & "") & "'," & _
'Val(rsTrans("NoteId") & "") & "," & rsTrans("Transaction_ID") & "," & Trans_DiscountType & "," & Val(Trans_Discount & "") & "," & _
'TaxValue & ",'" & order_no & "'," & SaleType & ",'" & Replace(cleanCashCustomerName, "'", "''") & "'," & _
'TaxAddValue & ",'" & CashCustomerPhone & "'," & SQLDate(rsTrans("last_changed"), True) & ","
'
'transSQL = transSQL & NetValue & "," & Transaction_NetValue & "," & IIf(DepandToConv, 1, 0) & "," & _
'CarTypeID & ",'" & PlateNo & "'," & OilsTypesID & "," & YearFact & ",'" & Shaseh & "','" & CarMeter & "'," & _
'FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Chasee & "'," & PPointID & ",'" & Phone2 & "'," & _
'SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & CarEnginoil & "," & _
'CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & CIBAN & "'," & SQLDate(RecTime, True) & ","
'
'transSQL = transSQL & SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
'InvoiceTypeCodeID & ",'" & InvoiceTypeCodename & "','" & DocumentCurrencyCode & "','" & _
'TaxCurrencyCode & "','" & Replace(paymentnote, "'", "''") & "','" & PaymentMeansCode & "')"
'
'transBatchSQL = transBatchSQL & transSQL & vbCrLf
'
'        '  ›«’Ì· Transaction
''        rsDetails.Open "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & rsTrans("Transaction_ID"), POSConnection
''        Do While Not rsDetails.EOF
''            detailsBatchSQL = detailsBatchSQL & "INSERT INTO Transaction_Details (Transaction_ID, Item_ID, Quantity, Price, SessionCode) VALUES (" & _
''                currentDestTransactionID & "," & rsDetails("Item_ID") & "," & rsDetails("Quantity") & "," & rsDetails("Price") & ",'" & SessionCode & "');"
''            rsDetails.MoveNext
''        Loop
''        rsDetails.Close
'
''Dim rsDetails As New ADODB.Recordset
'
'If recCounter = 0 Or recCounter Mod 10 = 0 Then
'    UpdateTransferCaption "Ã«—Ì ﬁ—«¡…  ›«’Ì· «·›Ê« Ì—", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'End If
'
'sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID"))
'rsDetails.CursorType = adOpenForwardOnly
'rsDetails.LockType = adLockReadOnly
'rsDetails.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'Do While Not rsDetails.EOF
'    Dim detailSQL As String
'
'    detailSQL = "INSERT INTO " & mServerD & "Transaction_Details (" & _
'        "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
'        "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
'        "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES ("
'
'    detailSQL = detailSQL & currentDestTransactionID & "," & Val(rsDetails("Item_ID")) & "," & Val(rsDetails("ItemCase") & "") & "," & _
'        Val(rsDetails("Quantity") & "") & "," & Val(rsDetails("Price") & "") & "," & Val(rsDetails("ItemDiscountType") & "") & "," & _
'        Val(rsDetails("ItemDiscount") & "") & "," & Val(rsDetails("ShowQty") & "") & "," & Val(rsDetails("showPrice") & "") & "," & _
'        Val(rsDetails("UnitId")) & "," & Val(rsDetails("ColorID") & "") & "," & Val(rsDetails("ItemSize") & "") & "," & _
'        Val(rsDetails("ClassId") & "") & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo") & "") & "," & _
'        Val(rsDetails("PumpId") & "") & "," & Val(rsDetails("PrevQty") & "") & ",'" & Replace(Trim(rsDetails("PrintName") & ""), "'", "''") & "'," & _
'        Val(rsDetails("Cash") & "") & "," & Val(rsDetails("Mada") & "") & "," & Val(rsDetails("Visa") & "") & "," & _
'        Val(rsDetails("Deferred") & "") & "," & Val(rsDetails("AmountH") & "") & "," & Val(rsDetails("AmountHComm") & "") & ","
'
'    detailSQL = detailSQL & "'" & Replace(Trim(rsDetails("DetailsPump") & ""), "'", "''") & "','" & Replace(Trim(rsDetails("Account_CodeComm") & ""), "'", "''") & "','" & _
'        Replace(Trim(rsDetails("Account_Code") & ""), "'", "''") & "'," & IIf(rsDetails("IsOther").Value, 1, 0) & ")"
'
'    detailsBatchSQL = detailsBatchSQL & detailSQL & vbCrLf
'
'    rsDetails.MoveNext
'Loop
'
'rsDetails.Close
'
'
'
''        ' ﬁÌ„… „÷«›…
''        rsValueAdded.Open "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & rsTrans("Transaction_ID"), POSConnection
''        Do While Not rsValueAdded.EOF
''            valueAddedBatchSQL = valueAddedBatchSQL & "INSERT INTO TransactionValueAdded (Transaction_ID, ItemID, VAT, Valu, SessionCode) VALUES (" & _
''                currentDestTransactionID & "," & rsValueAdded("ItemID") & "," & rsValueAdded("VAT") & "," & rsValueAdded("Valu") & ",'" & SessionCode & "');"
''            rsValueAdded.MoveNext
''        Loop
''        rsValueAdded.Close
'
'
''Dim rsValueAdded As New ADODB.Recordset
'If recCounter = 0 Or recCounter Mod 10 = 0 Then
'    UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·÷—Ì»… Ê«·ﬁÌ„… «·„÷«›…", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'End If
'sql = "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID"))
'rsValueAdded.CursorType = adOpenForwardOnly
'rsValueAdded.LockType = adLockReadOnly
'rsValueAdded.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'Do While Not rsValueAdded.EOF
'    Dim valueSQL As String
'    valueSQL = "INSERT INTO " & mServerD & "TransactionValueAdded (" & _
'                "Transaction_ID, ItemID, Vatyo, VAT, Valu, selectd, Transaction_Type, SessionCode) VALUES ("
'    valueSQL = valueSQL & currentDestTransactionID & "," & _
'        Val(rsValueAdded("ItemID")) & "," & _
'        Val(rsValueAdded("Vatyo")) & "," & _
'        Val(rsValueAdded("Vat")) & "," & _
'        Val(rsValueAdded("Valu")) & "," & _
'        Val(rsValueAdded("selectd")) & "," & _
'        Val(rsValueAdded("Transaction_Type")) & ",'" & SessionCode & "')"
'
'    valueAddedBatchSQL = valueAddedBatchSQL & valueSQL & vbCrLf
'    rsValueAdded.MoveNext
'Loop
'rsValueAdded.Close
'
'If recCounter = 0 Or recCounter Mod 10 = 0 Then
'    UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·„œ›Ê⁄« ", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'End If
'
'If Val(rsTrans("Transaction_Type")) = 21 Or Val(rsTrans("Transaction_Type")) = 9 Then
'    'Dim rsPayments As New ADODB.Recordset
'    sql = "SELECT * FROM TblTransactionPayments WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID"))
'    rsPayments.CursorType = adOpenForwardOnly
'    rsPayments.LockType = adLockReadOnly
'    rsPayments.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    Do While Not rsPayments.EOF
'        Dim paymentSQL As String, Recorddate As Date
'        If IsNull(rsPayments("Recorddate")) Or Trim(rsPayments("Recorddate") & "") = "" Then
'            Recorddate = Now
'        Else
'            Recorddate = rsPayments("Recorddate")
'        End If
'
'        paymentSQL = "INSERT INTO " & mServerD & "TblTransactionPayments (" & _
'            "Transaction_ID, boxid, Recorddate, PointID, CurrentCashireID, PaymentID, Value, CardNo, Effect, SessionCode) VALUES ("
'        paymentSQL = paymentSQL & currentDestTransactionID & "," & _
'            Val(rsPayments("boxid")) & "," & SQLDate(Recorddate, True) & "," & _
'            Val(rsPayments("PointID")) & "," & Val(rsPayments("CurrentCashireID")) & "," & Val(rsPayments("PaymentID")) & "," & _
'            Val(rsPayments("Value")) & ",'" & Replace(rsPayments("CardNo") & "", "'", "''") & "'," & _
'            Val(rsPayments("Effect")) & ",'" & SessionCode & "')"
'
'        paymentsBatchSQL = paymentsBatchSQL & paymentSQL & vbCrLf
'        rsPayments.MoveNext
'    Loop
'    rsPayments.Close
'End If
'
'
'If Val(rsTrans("Transaction_Type")) = 21 Or Val(rsTrans("Transaction_Type")) = 9 Then
'    Set rsPayments = New ADODB.Recordset
'    sql = "SELECT * FROM TblSalesPayment WHERE TransID = " & Val(rsTrans("Transaction_ID"))
'    rsPayments.CursorType = adOpenForwardOnly
'    rsPayments.LockType = adLockReadOnly
'    rsPayments.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    Do While Not rsPayments.EOF
'        Dim paymentSQL2 As String
'        paymentSQL2 = "INSERT INTO " & mServerD & "TblSalesPayment (" & _
'            "TransID, PaymentID, Value) VALUES (" & _
'            currentDestTransactionID & "," & Val(rsPayments("PaymentID")) & "," & Val(rsPayments("Value")) & ")"
'
'        paymentsBatchSQL2 = paymentsBatchSQL2 & paymentSQL2 & vbCrLf
'        rsPayments.MoveNext
'    Loop
'    rsPayments.Close
'End If
'
'        recCounter = recCounter + 1
'    If recCounter Mod BatchSize = 0 Then
'
'    BatchNo = BatchNo + 1
'UpdateTransferCaption "Ã«—Ì  —ÕÌ· «·œ›⁄… —ﬁ„ " & BatchNo & " ≈·Ï «·”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'WritePhaseLog "Executing batch", "BatchNo=" & BatchNo & ", RecCounter=" & recCounter & ", Total=" & TotalInvoices
'
'
''        If transBatchSQL <> "" Then
''            LastSQL = transBatchSQL
''            WriteLog "Executing transBatchSQL batch", transBatchSQL
''            Cn.Execute transBatchSQL
''        End If
''
''        If detailsBatchSQL <> "" Then
''            LastSQL = detailsBatchSQL
''            WriteLog "Executing detailsBatchSQL batch", detailsBatchSQL
''            Cn.Execute detailsBatchSQL
''        End If
''
''        If valueAddedBatchSQL <> "" Then
''            LastSQL = valueAddedBatchSQL
''            WriteLog "Executing valueAddedBatchSQL batch", valueAddedBatchSQL
''            Cn.Execute valueAddedBatchSQL
''        End If
''
''        If paymentsBatchSQL2 <> "" Then  ' TblSalesPayment
''            LastSQL = paymentsBatchSQL2
''            WriteLog "Executing paymentsBatchSQL2 batch", paymentsBatchSQL2
''            Cn.Execute paymentsBatchSQL2
''        End If
''
''        If paymentsBatchSQL <> "" Then   ' TblTransactionPayments
''            LastSQL = paymentsBatchSQL
''            WriteLog "Executing paymentsBatchSQL batch", paymentsBatchSQL
''            Cn.Execute paymentsBatchSQL
'        'End If
'
'        If transBatchSQL <> "" Then
'    UpdateTransferCaption "Ã«—Ì  —ÕÌ· —ƒÊ” «·›Ê« Ì— ··”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'    LastSQL = transBatchSQL
'    WriteLog "Executing transBatchSQL batch", transBatchSQL
'    Cn.Execute transBatchSQL
'End If
'
'If detailsBatchSQL <> "" Then
'    UpdateTransferCaption "Ã«—Ì  —ÕÌ·  ›«’Ì· «·›Ê« Ì— ··”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'    LastSQL = detailsBatchSQL
'    WriteLog "Executing detailsBatchSQL batch", detailsBatchSQL
'    Cn.Execute detailsBatchSQL
'End If
'
'If valueAddedBatchSQL <> "" Then
'    UpdateTransferCaption "Ã«—Ì  —ÕÌ· «·÷—Ì»… Ê«·ﬁÌ„… «·„÷«›…", recCounter, TotalInvoices, SessionCode, mTimeStart
'    LastSQL = valueAddedBatchSQL
'    WriteLog "Executing valueAddedBatchSQL batch", valueAddedBatchSQL
'    Cn.Execute valueAddedBatchSQL
'End If
'
'If paymentsBatchSQL2 <> "" Then
'    UpdateTransferCaption "Ã«—Ì  —ÕÌ· TblSalesPayment", recCounter, TotalInvoices, SessionCode, mTimeStart
'    LastSQL = paymentsBatchSQL2
'    WriteLog "Executing paymentsBatchSQL2 batch", paymentsBatchSQL2
'    Cn.Execute paymentsBatchSQL2
'End If
'
'If paymentsBatchSQL <> "" Then
'    UpdateTransferCaption "Ã«—Ì  —ÕÌ· TblTransactionPayments", recCounter, TotalInvoices, SessionCode, mTimeStart
'    LastSQL = paymentsBatchSQL
'    WriteLog "Executing paymentsBatchSQL batch", paymentsBatchSQL
'    Cn.Execute paymentsBatchSQL
'End If
'
'
'        transBatchSQL = "": detailsBatchSQL = "": valueAddedBatchSQL = "": paymentsBatchSQL2 = "": paymentsBatchSQL = ""
'    End If
'
'
'        rsTrans.MoveNext
'    Loop
'UpdateTransferCaption "Ã«—Ì  —ÕÌ· ¬Œ— œ›⁄… ≈·Ï «·”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'WritePhaseLog "Executing final batch", "RecCounter=" & recCounter & ", Total=" & TotalInvoices
'If transBatchSQL <> "" Then
'    LastSQL = transBatchSQL
'    WriteLog "Executing transBatchSQL final", transBatchSQL
'    Cn.Execute transBatchSQL
'End If
'
'If detailsBatchSQL <> "" Then
'    LastSQL = detailsBatchSQL
'    WriteLog "Executing detailsBatchSQL final", detailsBatchSQL
'    Cn.Execute detailsBatchSQL
'End If
'
'If valueAddedBatchSQL <> "" Then
'    LastSQL = valueAddedBatchSQL
'    WriteLog "Executing valueAddedBatchSQL final", valueAddedBatchSQL
'    Cn.Execute valueAddedBatchSQL
'End If
'
'If paymentsBatchSQL2 <> "" Then
'    LastSQL = paymentsBatchSQL2
'    WriteLog "Executing paymentsBatchSQL2 final", paymentsBatchSQL2
'    Cn.Execute paymentsBatchSQL2
'End If
'
'If paymentsBatchSQL <> "" Then
'    LastSQL = paymentsBatchSQL
'    WriteLog "Executing paymentsBatchSQL final", paymentsBatchSQL
'    Cn.Execute paymentsBatchSQL
'End If
'
'transBatchSQL = "": detailsBatchSQL = "": valueAddedBatchSQL = "": paymentsBatchSQL2 = "": paymentsBatchSQL = ""
'
'
'UpdateTransferCaption " „ ‰ﬁ· «·»Ì«‰« ° Ã«—Ì «·„—«Ã⁄… Ê«·„ÿ«»ﬁ…", recCounter, TotalInvoices, SessionCode, mTimeStart
'WritePhaseLog "Start reconcile/checksum", "SessionCode=" & SessionCode
''=== [2] POS source counters for this Session ===
'Dim SrcHeads As Long, SrcDet As Long, SrcVAT As Long, SrcPay As Long, SrcPay2 As Long
'
'
'' «·—ƒÊ”
'Set rsCnt = POSConnection.Execute( _
'  "SELECT COUNT(*) Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'" _
')
'SrcHeads = CLng(rsCnt!Cnt): rsCnt.Close
'
'' «· ›«’Ì·
'Set rsCnt = POSConnection.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM Transaction_Details d " & _
'  "WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'SrcDet = CLng(rsCnt!Cnt): rsCnt.Close
'
'' «·ﬁÌ„… «·„÷«›…
'Set rsCnt = POSConnection.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM TransactionValueAdded v " & _
'  "WHERE v.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'SrcVAT = CLng(rsCnt!Cnt): rsCnt.Close
'
'' TransactionPayments
'Set rsCnt = POSConnection.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM TblTransactionPayments p " & _
'  "WHERE p.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'SrcPay = CLng(rsCnt!Cnt): rsCnt.Close
'
'' SalesPayment
'Set rsCnt = POSConnection.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM TblSalesPayment s " & _
'  "WHERE s.TransID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'SrcPay2 = CLng(rsCnt!Cnt): rsCnt.Close
'Set rsCnt = Nothing
'
'UpdateTransferCaption "Ã«—Ì „—«Ã⁄… «·≈Ã„«·Ì«  Ê«·ﬂ„Ì« ", recCounter, TotalInvoices, SessionCode, mTimeStart
''=== [2.5] Checksums (POS vs Server) for this Session ===
'Dim srcQty As Double, dstQty As Double
'Dim SrcAmount As Currency, DstAmount As Currency
'Dim SrcVATSum As Currency, DstVATSum As Currency
'Dim SrcTPay As Currency, DstTPay As Currency
'Dim SrcSPay As Currency, DstSPay As Currency
'
''--- POS:  ›«’Ì· (ﬂ„Ì… + ﬁÌ„… «·»‰Êœ) ---
'Set rsCnt = POSConnection.Execute( _
'    "SELECT " & _
'    "  SUM(CAST(d.Quantity AS float)) AS SumQty, " & _
'    "  SUM(CAST(d.Quantity * d.Price AS decimal(18,4))) AS SumAmount " & _
'    "FROM Transaction_Details d " & _
'    "WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumQty) Then srcQty = 0# Else srcQty = CDbl(rsCnt!SumQty)
'    If IsNull(rsCnt!SumAmount) Then SrcAmount = CCur(0) Else SrcAmount = CCur(rsCnt!SumAmount)
'End If
'rsCnt.Close
'
''--- POS: «·÷—Ì»… ---
'Set rsCnt = POSConnection.Execute( _
'    "SELECT SUM(CAST(v.Valu AS decimal(18,4))) AS SumVAT " & _
'    "FROM TransactionValueAdded v " & _
'    "WHERE v.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumVAT) Then SrcVATSum = CCur(0) Else SrcVATSum = CCur(rsCnt!SumVAT)
'End If
'rsCnt.Close
'
''--- POS: «·„œ›Ê⁄«  (TblTransactionPayments) ---
'Set rsCnt = POSConnection.Execute( _
'    "SELECT SUM(CAST(p.Value AS decimal(18,4))) AS SumPay " & _
'    "FROM TblTransactionPayments p " & _
'    "WHERE p.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumPay) Then SrcTPay = CCur(0) Else SrcTPay = CCur(rsCnt!SumPay)
'End If
'rsCnt.Close
'
''--- POS: „œ›Ê⁄«  «·»Ì⁄ (TblSalesPayment) ---
'Set rsCnt = POSConnection.Execute( _
'    "SELECT SUM(CAST(s.Value AS decimal(18,4))) AS SumPay2 " & _
'    "FROM TblSalesPayment s " & _
'    "WHERE s.TransID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumPay2) Then SrcSPay = CCur(0) Else SrcSPay = CCur(rsCnt!SumPay2)
'End If
'rsCnt.Close
'Set rsCnt = Nothing
'
''--- Server:  ›«’Ì· (ﬂ„Ì… + ﬁÌ„… «·»‰Êœ) ---
'Set rsCnt = Cn.Execute( _
'    "SELECT " & _
'    "  SUM(CAST(d.Quantity AS float)) AS SumQty, " & _
'    "  SUM(CAST(d.Quantity * d.Price AS decimal(18,4))) AS SumAmount " & _
'    "FROM " & mServerD & "Transaction_Details d " & _
'    "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = d.Transaction_ID " & _
'    "WHERE t.SessionCode='" & SessionCode & "'" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumQty) Then dstQty = 0# Else dstQty = CDbl(rsCnt!SumQty)
'    If IsNull(rsCnt!SumAmount) Then DstAmount = CCur(0) Else DstAmount = CCur(rsCnt!SumAmount)
'End If
'rsCnt.Close
'
''--- Server: «·÷—Ì»… ---
'Set rsCnt = Cn.Execute( _
'    "SELECT SUM(CAST(v.Valu AS decimal(18,4))) AS SumVAT " & _
'    "FROM " & mServerD & "TransactionValueAdded v " & _
'    "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = v.Transaction_ID " & _
'    "WHERE t.SessionCode='" & SessionCode & "'" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumVAT) Then DstVATSum = CCur(0) Else DstVATSum = CCur(rsCnt!SumVAT)
'End If
'rsCnt.Close
'
''--- Server: «·„œ›Ê⁄«  (TblTransactionPayments) ---
'Set rsCnt = Cn.Execute( _
'    "SELECT SUM(CAST(p.Value AS decimal(18,4))) AS SumPay " & _
'    "FROM " & mServerD & "TblTransactionPayments p " & _
'    "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = p.Transaction_ID " & _
'    "WHERE t.SessionCode='" & SessionCode & "'" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumPay) Then DstTPay = CCur(0) Else DstTPay = CCur(rsCnt!SumPay)
'End If
'rsCnt.Close
'
''--- Server: „œ›Ê⁄«  «·»Ì⁄ (TblSalesPayment) ---
'Set rsCnt = Cn.Execute( _
'    "SELECT SUM(CAST(s.Value AS decimal(18,4))) AS SumPay2 " & _
'    "FROM " & mServerD & "TblSalesPayment s " & _
'    "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = s.TransID " & _
'    "WHERE t.SessionCode='" & SessionCode & "'" _
')
'If Not rsCnt.EOF Then
'    If IsNull(rsCnt!SumPay2) Then DstSPay = CCur(0) Else DstSPay = CCur(rsCnt!SumPay2)
'End If
'rsCnt.Close
'Set rsCnt = Nothing
'
''--- Â«„‘ «·”„«Õ ---
'Dim epsQty As Double:    epsQty = 0.0001   ' ··ﬂ„Ì« 
'Dim epsMoney As Currency: epsMoney = 0.01  ' 1 ﬁ—‘/Â··…
'
''--- „ﬁ«—‰… Checksums ---
'If (Abs(srcQty - dstQty) > epsQty) _
'   Or (Abs(SrcAmount - DstAmount) > epsMoney) _
'   Or (Abs(SrcVATSum - DstVATSum) > epsMoney) _
'   Or (Abs(SrcTPay - DstTPay) > epsMoney) _
'   Or (Abs(SrcSPay - DstSPay) > epsMoney) Then
'
'    WriteLog "Checksum mismatch: " & _
'             "Qty " & FormatNumber(srcQty, 6) & "/" & FormatNumber(dstQty, 6) & _
'             ", Amount " & FormatCurrency(SrcAmount) & "/" & FormatCurrency(DstAmount) & _
'             ", VAT " & FormatCurrency(SrcVATSum) & "/" & FormatCurrency(DstVATSum) & _
'             ", Pay " & FormatCurrency(SrcTPay) & "/" & FormatCurrency(DstTPay) & _
'             ", Pay2 " & FormatCurrency(SrcSPay) & "/" & FormatCurrency(DstSPay), ""
'
'    Err.Raise vbObjectError + 779, , "Checksum reconcile failed for SessionCode=" & SessionCode
'End If
'
'
''=== [3] Server destination counters for this Session ===
'Dim DstHeads As Long, DstDet As Long, DstVAT As Long, DstPay As Long, DstPay2 As Long
'
'' «·—ƒÊ”
'Set rsCnt = Cn.Execute( _
'  "SELECT COUNT(*) Cnt FROM " & mServerD & "Transactions WHERE SessionCode='" & SessionCode & "'" _
')
'DstHeads = CLng(rsCnt!Cnt): rsCnt.Close
'
'' «· ›«’Ì·
'Set rsCnt = Cn.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM " & mServerD & "Transaction_Details d " & _
'  "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=d.Transaction_ID " & _
'  "WHERE t.SessionCode='" & SessionCode & "'" _
')
'DstDet = CLng(rsCnt!Cnt): rsCnt.Close
'
'' «·ﬁÌ„… «·„÷«›…
'Set rsCnt = Cn.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM " & mServerD & "TransactionValueAdded v " & _
'  "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=v.Transaction_ID " & _
'  "WHERE t.SessionCode='" & SessionCode & "'" _
')
'DstVAT = CLng(rsCnt!Cnt): rsCnt.Close
'
'' TransactionPayments
'Set rsCnt = Cn.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM " & mServerD & "TblTransactionPayments p " & _
'  "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=p.Transaction_ID " & _
'  "WHERE t.SessionCode='" & SessionCode & "'" _
')
'DstPay = CLng(rsCnt!Cnt): rsCnt.Close
'
'' SalesPayment
'Set rsCnt = Cn.Execute( _
'  "SELECT COUNT(*) Cnt " & _
'  "FROM " & mServerD & "TblSalesPayment s " & _
'  "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=s.TransID " & _
'  "WHERE t.SessionCode='" & SessionCode & "'" _
')
'DstPay2 = CLng(rsCnt!Cnt): rsCnt.Close
'Set rsCnt = Nothing
'UpdateTransferCaption "Ã«—Ì «· Õﬁﬁ «·‰Â«∆Ì ﬁ»· «⁄ „«œ «·‰ﬁ·", recCounter, TotalInvoices, SessionCode, mTimeStart
'
''=== [4] Reconcile before marking Copied ===
'If (DstHeads <> SrcHeads) Or (DstDet <> SrcDet) Or (DstVAT <> SrcVAT) Or (DstPay <> SrcPay) Or (DstPay2 <> SrcPay2) Then
'    WriteLog "Reconcile mismatch: Heads " & SrcHeads & "/" & DstHeads & _
'             ", Det " & SrcDet & "/" & DstDet & _
'             ", VAT " & SrcVAT & "/" & DstVAT & _
'             ", Pay " & SrcPay & "/" & DstPay & _
'             ", Pay2 " & SrcPay2 & "/" & DstPay2, ""
'    Err.Raise vbObjectError + 778, , "Reconcile failed for SessionCode=" & SessionCode
'End If
'
'
''=========================================
'' Consistency check BEFORE marking Copied
''=========================================
'Set rsCnt = New ADODB.Recordset
'Dim srcCount As Long, dstCount As Long
'
'' ⁄œœ «·›Ê« Ì— «·„Ê”Ê„… ›Ì «·‰ﬁÿ… ·Â–Â «·Ã·”… («·„›—Ê÷  ﬂÊ‰ ÂÌ „Ã„Ê⁄… «·⁄„·)
'Set rsCnt = POSConnection.Execute( _
'    "SELECT COUNT(*) AS Cnt " & _
'    "FROM Transactions " & _
'    "WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "'" _
')
'srcCount = CLng(rsCnt.Fields("Cnt").Value)
'rsCnt.Close
'Set rsCnt = Nothing
'
'' ⁄œœ —ƒÊ” «·›Ê« Ì— «··Ì « œ—Ã  ›⁄·« ⁄·Ï «·”Ì—›— »Â–Â «·Ã·”…
'Set rsCnt = Cn.Execute( _
'    "SELECT COUNT(*) AS Cnt " & _
'    "FROM " & mServerD & "Transactions " & _
'    "WHERE SessionCode = '" & SessionCode & "'" _
')
'dstCount = CLng(rsCnt.Fields("Cnt").Value)
'rsCnt.Close
'Set rsCnt = Nothing
'
'
'If dstCount <> srcCount Then
'    WriteLog "Session mismatch before marking Copied. src=" & srcCount & " dst=" & dstCount, ""
'    Err.Raise vbObjectError + 777, , "Mismatch between source and server counts for SessionCode=" & SessionCode
'End If
'
'UpdateTransferCaption " „ «· Õﬁﬁ° Ã«—Ì «⁄ „«œ «·›Ê« Ì— ﬂ„ı‰ﬁÊ·…", recCounter, TotalInvoices, SessionCode, mTimeStart
'
'POSConnection.Execute _
'    "UPDATE Transactions " & _
'    "SET Copied = 1,SessionCode=NULL " & _
'    "WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "'"
'
'    ' ⁄·«„… «·‰ﬁ· Ê«·‹Commit
'    'POSConnection.Execute "UPDATE Transactions SET Copied = 1, SessionCode = '" & SessionCode & "' WHERE IsNull(Copied,0)=0 AND " & GetQuery
'
'    Cn.CommitTrans
'
'UpdateTransferCaption " „ «·‰ﬁ· »‰Ã«Õ", TotalInvoices, TotalInvoices, SessionCode, mTimeStart
'WritePhaseLog "Transfer completed successfully", "SessionCode=" & SessionCode & ", Count=" & TotalInvoices
'MousePointer = vbDefault
'
'
'Dim r As Integer
'r = 1
'
'' 1) ⁄œœ «·›Ê« Ì— («·—ƒÊ”)
'grdInfo.TextMatrix(r, 0) = "⁄œœ «·›Ê« Ì—"
'grdInfo.TextMatrix(r, 1) = CStr(SrcHeads)
'grdInfo.TextMatrix(r, 2) = CStr(DstHeads)
'grdInfo.TextMatrix(r, 3) = IIf(SrcHeads = DstHeads, "?", "?")
'r = r + 1
'
'' 2) ⁄œœ ”ÿÊ— «· ›«’Ì·
'grdInfo.TextMatrix(r, 0) = "⁄œœ ”ÿÊ— «· ›«’Ì·"
'grdInfo.TextMatrix(r, 1) = CStr(SrcDet)
'grdInfo.TextMatrix(r, 2) = CStr(DstDet)
'grdInfo.TextMatrix(r, 3) = IIf(SrcDet = DstDet, "?", "?")
'r = r + 1
'
''' 3) ⁄œœ ”ÿÊ— «·ﬁÌ„… «·„÷«›…
'grdInfo.TextMatrix(r, 0) = "«Ã„«·Ì ﬁÌ„ «·›Ê« Ì—"
'grdInfo.TextMatrix(r, 1) = CStr(SrcAmount)
'grdInfo.TextMatrix(r, 2) = CStr(DstAmount)
'grdInfo.TextMatrix(r, 3) = IIf(SrcAmount = DstAmount, "0", "ÌÊÃœ „‘ﬂ·… ›Ï «·«Ã„«·Ì« ")
'r = r + 1
'
'' 4) ⁄œœ ”Ã·«  TblTransactionPayments
'grdInfo.TextMatrix(r, 0) = "⁄œœ ”Ã·«  TransactionPayments"
'grdInfo.TextMatrix(r, 1) = CStr(SrcPay)
'grdInfo.TextMatrix(r, 2) = CStr(DstPay)
'grdInfo.TextMatrix(r, 3) = IIf(SrcPay = DstPay, "?", "?")
'r = r + 1
'
'' 5) ⁄œœ ”Ã·«  TblSalesPayment
'grdInfo.TextMatrix(r, 0) = "⁄œœ ”Ã·«  SalesPayment"
'grdInfo.TextMatrix(r, 1) = CStr(SrcPay2)
'grdInfo.TextMatrix(r, 2) = CStr(DstPay2)
'grdInfo.TextMatrix(r, 3) = IIf(SrcPay2 = DstPay2, "?", "?")
'
'' («Œ Ì«—Ì)  ‰”Ìﬁ«  »”Ìÿ…
' 'grdiInfo.ColWidth(0) = 2200: grdInfo.ColWidth(1) = 1600: grdInfo.ColWidth(2) = 1600: grdInfo.ColWidth(3) = 900
' grdInfo.ColAlignment(1) = 7: grdInfo.ColAlignment(2) = 7: grdInfo.ColAlignment(3) = 4
'
'Dim elapsedSec As Long, elapsedMin As Long
'
'elapsedSec = DateDiff("s", mTimeStart, Now)
'elapsedMin = elapsedSec \ 60
'elapsedSec = elapsedSec Mod 60
'frmPopup.ShowMessage " „ «·‰ﬁ· »‰Ã«Õ." & vbCrLf & _
'       "«·Êﬁ  «·„” €—ﬁ: " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì….", vbInformation
'
'lblWait.Caption = " „ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄«  »‰Ã«Õ." & _
'                  " «·Êﬁ  «·„” €—ﬁ: " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì…."
'txtEndTime = CStr(Now)
'
'
'
''=== Õ”«» «·“„‰ ===
'
'elapsedSec = DateDiff("s", mTimeStart, Now)
'
''===  ÕœÌœ «·« Ã«Â Ê‰Ê⁄ «·‰ﬁ· ===
'Dim direction As String, kind As String
'direction = "POS->Server"        ' √Ê "Server->POS" ›Ì ≈Ã—«¡ «· ÕÊÌ·«  «·„Œ“‰Ì…
'kind = "Sales"                   ' √Ê "Transfers"
'
''=== „À«· ··ﬁÌ„ («” ⁄„· „ €Ì—« ﬂ «·„Õ”Ê»… ›⁄·«) ===
'Call SaveSyncLog( _
'    POSConnection, Cn, _
'    SessionCode, direction, kind, _
'    mTimeStart, Now, elapsedSec, _
'    POSlServer, POSDb, SysSQLServerName, ServerDb, _
'    BranchID, GetQuery, _
'    BatchSize, FetchSize, _
'    SrcHeads, DstHeads, _
'    SrcDet, DstDet, _
'    SrcVAT, DstVAT, _
'    SrcPay, DstPay, _
'    SrcPay2, DstPay2, _
'    SrcAmount, DstAmount, _
'    SrcVATSum, DstVATSum, _
'    SrcTPay, DstTPay, _
'    SrcSPay, DstSPay, _
'    True, "" _
')
'
'
'    POSname_Change
'   ' MsgBox " „ «·‰ﬁ· »‰Ã«Õ"
'
'EndSub:
'    'lblWait.Visible = False
'    Exit Sub
'
'ErrorHandler:
'    Cn.RollbackTrans
'    POSConnection.Execute "UPDATE Transactions SET Copied = null, SessionCode = null WHERE SessionCode = '" & SessionCode & "'"
'    WriteLog "ErrorHandler: " & Err.Description, LastSQL
'
'    ' œ«Œ· ErrorHandler ﬁ»· «·—”«·…:
'Call SaveSyncLog(POSConnection, Cn, SessionCode, direction, kind, _
'                 mTimeStart, Now, DateDiff("s", mTimeStart, Now), _
'                 POSlServer, POSDb, SysSQLServerName, ServerDb, _
'                 BranchID, GetQuery, BatchSize, FetchSize, _
'                 SrcHeads, DstHeads, SrcDet, DstDet, SrcVAT, DstVAT, _
'                 SrcPay, DstPay, SrcPay2, DstPay2, _
'                 SrcAmount, DstAmount, SrcVATSum, DstVATSum, _
'                 SrcTPay, DstTPay, SrcSPay, DstSPay, _
'                 False, Err.Description)
'
'
'    frmPopup.ShowMessage "Œÿ√ «À‰«¡ «·‰ﬁ· —Ã«¡ «· Ê«’· „⁄ „”∆Ê·Ï «·‰Ÿ«„: " & Err.Description, vbCritical
'    lblWait.Visible = False
'End Sub
'
'



'Private Sub Command14_Click()
'
'    On Error GoTo ErrorHandler
'
'    '========================
'    ' Declarations
'    '========================
'    Dim POSCn As ADODB.Connection
'    Dim rsCnt As ADODB.Recordset
'    Dim rsTrans As ADODB.Recordset
'    Dim rsDetails As ADODB.Recordset
'    Dim rsValueAdded As ADODB.Recordset
'    Dim rsPayments As ADODB.Recordset
'    Dim rsSer As ADODB.Recordset
'
'    Dim BatchSize As Integer
'    Dim recCounter As Long
'    Dim recCount As Long
'    Dim BatchNo As Long
'    Dim TotalInvoices As Long
'
'    Dim transBatchSQL As String
'    Dim detailsBatchSQL As String
'    Dim valueAddedBatchSQL As String
'    Dim paymentsBatchSQL2 As String
'    Dim paymentsBatchSQL As String
'    Dim LastSQL As String
'    Dim sql As String
'    Dim transSQL As String
'
'    Dim SessionCode As String
'    Dim CurrentInvoiceNo As String
'
'    Dim mTimeStart As Date
'    Dim elapsedSec As Long, elapsedMin As Long
'
'    Dim direction As String, kind As String
'    Dim FetchSize As Long
'
'    Dim CarOilChangeDate As Date, RecTime As Date, mTimeIn As String
'    Dim ActualDeliveryDate As Date, LatestDeliveryDate As Date
'    Dim FromTransaction_Date As Date
'    Dim FromTransaction_Type As Long
'    Dim FromTransaction_ID As Double
'
'    Dim PayMentType As Long, cusID As Long, BranchID As Integer, BoxID As Long, BillBasedOn As Double
'    Dim VAT As Double, VATYou As Double, NoteId As Long, Trans_DiscountType As Long
'    Dim Trans_Discount As Double, TaxValue As Double, order_no As Long, SaleType As Long
'    Dim TaxAddValue As Double, NetValue As Double, Transaction_NetValue As Double, DepandToConv As Long
'    Dim CarTypeID As Long, OilsTypesID As Long, YearFact As Long, FixesAssetsID As Long, ColorID2 As Long
'    Dim KM As Double, PPointID As Long, SupplerID As Long, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double
'    Dim CarEnginoil As Double, CarGearOil As Double, InvoiceTypeCodeID As Long
'    Dim storeID As Variant, userID As Variant, Emp_ID As Variant
'    Dim NoteSerial As String, NoteSerial1 As String, TransactionComment As String
'    Dim CashCustomerName As String, CashCustomerPhone As String
'    Dim cleanCashCustomerName As String
'    Dim PlateNo As String, Shaseh As String, CarMeter As String
'    Dim Chasee As String, Phone2 As String
'    Dim CIBAN As String
'    Dim InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String
'    Dim paymentnote As String, PaymentMeansCode As String
'
'    ' Counters for reconcile
'    Dim SrcHeads As Long, SrcDet As Long, SrcVAT As Long, SrcPay As Long, SrcPay2 As Long
'    Dim DstHeads As Long, DstDet As Long, DstVAT As Long, DstPay As Long, DstPay2 As Long
'    Dim srcCount As Long, dstCount As Long
'
'    ' Checksums
'    Dim srcQty As Double, dstQty As Double
'    Dim SrcAmount As Currency, DstAmount As Currency
'    Dim SrcVATSum As Currency, DstVATSum As Currency
'    Dim SrcTPay As Currency, DstTPay As Currency
'    Dim SrcSPay As Currency, DstSPay As Currency
'    Dim epsQty As Double
'    Dim epsMoney As Currency
'
'    ' Misc
'    Dim currentDestTransactionID As String
'    Dim tmpRecTime As Variant
'    Dim v As Variant
'    Dim r As Integer
'
'    '========================
'    ' Initial values
'    '========================
'    BatchSize = 50
'    recCounter = 0
'    recCount = 0
'    BatchNo = 0
'    FetchSize = 0
'
'    transBatchSQL = ""
'    detailsBatchSQL = ""
'    valueAddedBatchSQL = ""
'    paymentsBatchSQL2 = ""
'    paymentsBatchSQL = ""
'    LastSQL = ""
'
'    direction = "POS->Server"
'    kind = "Sales"
'
'    mTimeStart = Now
'    txtStartTime = mTimeStart
'
'    '========================
'    ' Validation
'    '========================
'    If POSlServer.Text = "" Then
'        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
'        Exit Sub
'    End If
'
'    If ConnectionFirst = False Then Exit Sub
'
'    lblWait.Visible = True
'    lblWait.Caption = "Ã«—Ì »œ¡ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄« ..."
'    DoEvents
'    MousePointer = vbHourglass
'    WritePhaseLog "Start Command14_Click"
'
'    SessionCode = Format(Now, "yyyymmddhhmmss")
'    UpdateTransferCaption "Ã«—Ì  ÃÂÌ“ Ã·”… «·‰ﬁ·", 0, 0, SessionCode, mTimeStart
'
'    '========================
'    ' Open local POS connection
'    '========================
'    Set POSCn = New ADODB.Connection
'    POSCn.CursorLocation = adUseServer
'    POSCn.ConnectionTimeout = 5000
'    POSCn.CommandTimeout = 5000
'    POSCn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'                             ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                             ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer.Text
'    POSCn.Open
'
'    ' Central connection already prepared by ConnectionFirst
'    Cn.CursorLocation = adUseServer
'
'    Text3 = "Query: " & GetQuery
'
'    '========================
'    ' Tag source transactions
'    '========================
'    LastSQL = "UPDATE Transactions SET SessionCode = '" & SessionCode & "' WHERE IsNull(Copied,0)=0 AND " & GetQuery
'    POSCn.Execute LastSQL
'
'    WritePhaseLog "Tagging source transactions", "SessionCode=" & SessionCode
'    UpdateTransferCaption " „  ⁄·Ì„ «·›Ê« Ì— «·„—«œ ‰ﬁ·Â«", 0, 0, SessionCode, mTimeStart
'
'    '========================
'    ' Get total count first
'    '========================
'    Set rsCnt = POSCn.Execute( _
'        "SELECT COUNT(*) AS Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND " & GetQuery)
'    TotalInvoices = CLng(rsCnt!Cnt)
'    rsCnt.Close
'    Set rsCnt = Nothing
'
'    UpdateTransferCaption " „ «·⁄ÀÊ— ⁄·Ï ›Ê« Ì— ··‰ﬁ·", 0, TotalInvoices, SessionCode, mTimeStart
'    WritePhaseLog "Invoices selected", "Count=" & TotalInvoices
'
'    If TotalInvoices = 0 Then
'        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
'        GoTo EndSub
'    End If
'
'    '========================
'    ' Open source transactions
'    '========================
'    Set rsTrans = New ADODB.Recordset
'    LastSQL = "SELECT * FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND " & GetQuery & " ORDER BY Transaction_ID"
'    rsTrans.Open LastSQL, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If rsTrans.EOF Then
'        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
'        GoTo EndSub
'    End If
'
'    '========================
'    ' Start SQL transaction on central server
'    '========================
'    LastSQL = "SET XACT_ABORT ON;"
'    Cn.Execute LastSQL
'    Cn.BeginTrans
'
'    '========================
'    ' Main Loop
'    '========================
'    Do While Not rsTrans.EOF
'
'        If Trim(rsTrans("NoteSerial1").Value & "") <> "" Then
'            CurrentInvoiceNo = Trim(rsTrans("NoteSerial1").Value & "")
'        Else
'            CurrentInvoiceNo = CStr(Val(rsTrans("Transaction_ID").Value & ""))
'        End If
'
'        If recCounter = 0 Or recCounter Mod 5 = 0 Then
'            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·›« Ê—… —ﬁ„ " & CurrentInvoiceNo, recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'        End If
'
'        recCount = recCount + 1
'
'        '========================
'        ' Read all needed values
'        '========================
'        PayMentType = Val(rsTrans("PaymentType").Value & "")
'        FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
'        cusID = Val(rsTrans("CusID").Value & "")
'        storeID = Val(rsTrans("StoreID").Value & "")
'        userID = Val(rsTrans("UserID").Value & "")
'        Emp_ID = Val(rsTrans("Emp_ID").Value & "")
'        BranchID = Val(rsTrans("BranchID").Value & "")
'        BoxID = Val(rsTrans("BoxID").Value & "")
'        BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
'        mTimeIn = Trim(rsTrans("TimeIn").Value & "")
'        VAT = Val(rsTrans("VAT").Value & "")
'        VATYou = Val(rsTrans("VATYou").Value & "")
'        NoteSerial = rsTrans("NoteSerial").Value & ""
'        NoteSerial1 = rsTrans("NoteSerial1").Value & ""
'        NoteId = Val(rsTrans("NoteId").Value & "")
'        TransactionComment = rsTrans("TransactionComment").Value & ""
'        Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
'        FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
'        Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
'        TaxValue = Val(rsTrans("TaxValue").Value & "")
'        order_no = Val(rsTrans("order_no").Value & "")
'        SaleType = Val(rsTrans("SaleType").Value & "")
'        CashCustomerName = rsTrans("CashCustomerName").Value & ""
'        cleanCashCustomerName = CashCustomerName
'        TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
'        CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
'        NetValue = Val(rsTrans("NetValue").Value & "")
'        Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
'        DepandToConv = Val(rsTrans("DepandToConv").Value & "")
'        CarTypeID = Val(rsTrans("CarTypeID").Value & "")
'        PlateNo = rsTrans("PlateNo").Value & ""
'        OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
'        YearFact = Val(rsTrans("YearFact").Value & "")
'        Shaseh = rsTrans("Shaseh").Value & ""
'        CarMeter = rsTrans("CarMeter").Value & ""
'        FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
'        ColorID2 = Val(rsTrans("ColorID2").Value & "")
'        KM = Val(rsTrans("KM").Value & "")
'        Chasee = rsTrans("Chasee").Value & ""
'        PPointID = Val(rsTrans("PPointID").Value & "")
'        Phone2 = rsTrans("Phone2").Value & ""
'        SupplerID = Val(rsTrans("SupplerID").Value & "")
'        Ser = Val(rsTrans("Ser").Value & "")
'        CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
'        CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
'        CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
'        CarGearOil = Val(rsTrans("CarGearOil").Value & "")
'
'        If Trim(rsTrans("CarOilChangeDate").Value & "") = "" Then
'            CarOilChangeDate = Date
'        Else
'            CarOilChangeDate = rsTrans("CarOilChangeDate").Value
'        End If
'
'        CIBAN = rsTrans("CIBAN").Value & ""
'
'        tmpRecTime = rsTrans("RecTime").Value
'        v = tmpRecTime
'        If IsDate(v) Then
'            If Year(CDate(v)) = 1899 And Month(CDate(v)) = 12 And Day(CDate(v)) = 30 Then
'                RecTime = Time
'            Else
'                RecTime = CDate(v)
'            End If
'        Else
'            RecTime = Time
'        End If
'
'        If Trim(rsTrans("ActualDeliveryDate").Value & "") = "" Then
'            ActualDeliveryDate = Date
'        Else
'            ActualDeliveryDate = rsTrans("ActualDeliveryDate").Value
'        End If
'
'        If Trim(rsTrans("LatestDeliveryDate").Value & "") = "" Then
'            LatestDeliveryDate = ActualDeliveryDate
'        Else
'            LatestDeliveryDate = rsTrans("LatestDeliveryDate").Value
'        End If
'
'        InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
'        InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
'        DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
'        TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
'        paymentnote = rsTrans("paymentnote").Value & ""
'        PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""
'        FromTransaction_Date = rsTrans("Transaction_Date").Value
'
'        ' ≈–« POSBillType = 0°  ⁄œÌ· NoteSerial Ê NoteId
'        If Val(rsTrans("POSBillType").Value & "") = 0 Then
'            NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
'            NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
'        End If
'
'        TransactionComment = " ›« Ê—… „‰ﬁÊ·… „‰ " & POSname.Text & "   " & _
'                             "   —ﬁ„ «·›« Ê—… " & NoteSerial1
'
'        '========================
'        ' Reserve destination Transaction_ID
'        '========================
'        LastSQL = "EXEC dbo.ReserveTransactionId"
'        If recCounter = 0 Or recCounter Mod 10 = 0 Then
'            UpdateTransferCaption "Ã«—Ì ÕÃ“ √—ﬁ«„ «·„⁄«„·«  ⁄·Ï «·”Ì—›—", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'        End If
'
'        Set rsSer = Cn.Execute(LastSQL)
'        If Not rsSer.EOF Then
'            currentDestTransactionID = CStr(rsSer.Fields("NewId").Value)
'        Else
'            Err.Raise vbObjectError + 500, , "·„ Ì „ ≈—Ã«⁄ Transaction_ID ÃœÌœ „‰ «·”Ì—›—"
'        End If
'        rsSer.Close
'        Set rsSer = Nothing
'
'        '========================
'        ' Build Transactions INSERT
'        '========================
'        transSQL = "INSERT INTO " & mServerD & "Transactions (" & _
'        "Transaction_ID, Transaction_Date,TimeIn ,TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
'        "CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, NoteSerial1, " & _
'        "NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, OldNoteserial, OldNoteId, " & _
'        "OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, SaleType, CashCustomerName, " & _
'        "TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, DepandToConv, CarTypeID, " & _
'        "PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, ColorID2, KM, Chasee, PPointID, Phone2, " & _
'        "SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, CarGearOil, CarOilChangeDate, CIBAN, RecTime, " & _
'        "ActualDeliveryDate, LatestDeliveryDate, InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, " & _
'        "TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("
'
'        transSQL = transSQL & currentDestTransactionID & "," & SQLDate(FromTransaction_Date, True) & ",'" & Trim(mTimeIn) & "'," & _
'        Val(rsTrans("TypeInvoice").Value & "") & ",'" & Replace(rsTrans("Transaction_Serial").Value & "", "'", "''") & "'," & _
'        FromTransaction_Type & "," & PayMentType & "," & cusID & "," & storeID & "," & userID & "," & _
'        Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & ",'" & _
'        Replace(NoteSerial, "'", "''") & "','" & Replace(NoteSerial1, "'", "''") & "'," & NoteId & ",1,'" & Replace(TransactionComment, "'", "''") & "','" & _
'        SessionCode & "'," & IIf(Val(rsTrans("POSBillType").Value & "") = 0, 1, Val(rsTrans("POSBillType").Value & "")) & ",'" & Replace(rsTrans("NoteSerial1").Value & "", "'", "''") & "','" & Replace(Trim(rsTrans("NoteSerial").Value & ""), "'", "''") & "'," & _
'        Val(rsTrans("NoteId").Value & "") & "," & Val(rsTrans("Transaction_ID").Value & "") & "," & Trans_DiscountType & "," & Val(Trans_Discount & "") & "," & _
'        TaxValue & ",'" & order_no & "'," & SaleType & ",'" & Replace(cleanCashCustomerName, "'", "''") & "'," & _
'        TaxAddValue & ",'" & Replace(CashCustomerPhone, "'", "''") & "'," & SQLDate(rsTrans("last_changed").Value, True) & ","
'
'        transSQL = transSQL & NetValue & "," & Transaction_NetValue & "," & IIf(Val(DepandToConv & "") <> 0, 1, 0) & "," & _
'        CarTypeID & ",'" & Replace(PlateNo, "'", "''") & "'," & OilsTypesID & "," & YearFact & ",'" & Replace(Shaseh, "'", "''") & "','" & Replace(CarMeter, "'", "''") & "'," & _
'        FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Replace(Chasee, "'", "''") & "'," & PPointID & ",'" & Replace(Phone2, "'", "''") & "'," & _
'        SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & CarEnginoil & "," & _
'        CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & Replace(CIBAN, "'", "''") & "'," & SQLDate(RecTime, True) & ","
'
'        transSQL = transSQL & SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
'        InvoiceTypeCodeID & ",'" & Replace(InvoiceTypeCodename, "'", "''") & "','" & Replace(DocumentCurrencyCode, "'", "''") & "','" & _
'        Replace(TaxCurrencyCode, "'", "''") & "','" & Replace(paymentnote, "'", "''") & "','" & Replace(PaymentMeansCode, "'", "''") & "')"
'
'        transBatchSQL = transBatchSQL & transSQL & vbCrLf
'
'        '========================
'        ' Transaction_Details
'        '========================
'        If recCounter = 0 Or recCounter Mod 10 = 0 Then
'            UpdateTransferCaption "Ã«—Ì ﬁ—«¡…  ›«’Ì· «·›Ê« Ì—", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'        End If
'
'        Set rsDetails = New ADODB.Recordset
'        sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
'        rsDetails.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'        Do While Not rsDetails.EOF
'            Dim detailSQL As String
'
'            detailSQL = "INSERT INTO " & mServerD & "Transaction_Details (" & _
'                "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
'                "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
'                "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES ("
'
'            detailSQL = detailSQL & currentDestTransactionID & "," & Val(rsDetails("Item_ID").Value & "") & "," & Val(rsDetails("ItemCase").Value & "") & "," & _
'                Val(rsDetails("Quantity").Value & "") & "," & Val(rsDetails("Price").Value & "") & "," & Val(rsDetails("ItemDiscountType").Value & "") & "," & _
'                Val(rsDetails("ItemDiscount").Value & "") & "," & Val(rsDetails("ShowQty").Value & "") & "," & Val(rsDetails("showPrice").Value & "") & "," & _
'                Val(rsDetails("UnitId").Value & "") & "," & Val(rsDetails("ColorID").Value & "") & "," & Val(rsDetails("ItemSize").Value & "") & "," & _
'                Val(rsDetails("ClassId").Value & "") & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo").Value & "") & "," & _
'                Val(rsDetails("PumpId").Value & "") & "," & Val(rsDetails("PrevQty").Value & "") & ",'" & Replace(Trim(rsDetails("PrintName").Value & ""), "'", "''") & "'," & _
'                Val(rsDetails("Cash").Value & "") & "," & Val(rsDetails("Mada").Value & "") & "," & Val(rsDetails("Visa").Value & "") & "," & _
'                Val(rsDetails("Deferred").Value & "") & "," & Val(rsDetails("AmountH").Value & "") & "," & Val(rsDetails("AmountHComm").Value & "") & ","
'
'            detailSQL = detailSQL & "'" & Replace(Trim(rsDetails("DetailsPump").Value & ""), "'", "''") & "','" & Replace(Trim(rsDetails("Account_CodeComm").Value & ""), "'", "''") & "','" & _
'                Replace(Trim(rsDetails("Account_Code").Value & ""), "'", "''") & "'," & IIf(Val(rsDetails("IsOther").Value & "") <> 0, 1, 0) & ")"
'
'            detailsBatchSQL = detailsBatchSQL & detailSQL & vbCrLf
'            rsDetails.MoveNext
'        Loop
'
'        rsDetails.Close
'        Set rsDetails = Nothing
'
'        '========================
'        ' TransactionValueAdded
'        '========================
'        If recCounter = 0 Or recCounter Mod 10 = 0 Then
'            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·÷—Ì»… Ê«·ﬁÌ„… «·„÷«›…", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'        End If
'
'        Set rsValueAdded = New ADODB.Recordset
'        sql = "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
'        rsValueAdded.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'        Do While Not rsValueAdded.EOF
'            Dim valueSQL As String
'
'            valueSQL = "INSERT INTO " & mServerD & "TransactionValueAdded (" & _
'                        "Transaction_ID, ItemID, Vatyo, VAT, Valu, selectd, Transaction_Type, SessionCode) VALUES ("
'            valueSQL = valueSQL & currentDestTransactionID & "," & _
'                Val(rsValueAdded("ItemID").Value & "") & "," & _
'                Val(rsValueAdded("Vatyo").Value & "") & "," & _
'                Val(rsValueAdded("Vat").Value & "") & "," & _
'                Val(rsValueAdded("Valu").Value & "") & "," & _
'                Val(rsValueAdded("selectd").Value & "") & "," & _
'                Val(rsValueAdded("Transaction_Type").Value & "") & ",'" & SessionCode & "')"
'
'            valueAddedBatchSQL = valueAddedBatchSQL & valueSQL & vbCrLf
'            rsValueAdded.MoveNext
'        Loop
'
'        rsValueAdded.Close
'        Set rsValueAdded = Nothing
'
'        '========================
'        ' Payments
'        '========================
'        If recCounter = 0 Or recCounter Mod 10 = 0 Then
'            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·„œ›Ê⁄« ", recCounter + 1, TotalInvoices, SessionCode, mTimeStart
'        End If
'
'        If Val(rsTrans("Transaction_Type").Value & "") = 21 Or Val(rsTrans("Transaction_Type").Value & "") = 9 Then
'
'            Set rsPayments = New ADODB.Recordset
'            sql = "SELECT * FROM TblTransactionPayments WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
'            rsPayments.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'            Do While Not rsPayments.EOF
'                Dim paymentSQL As String, Recorddate As Date
'
'                If IsNull(rsPayments("Recorddate").Value) Or Trim(rsPayments("Recorddate").Value & "") = "" Then
'                    Recorddate = Now
'                Else
'                    Recorddate = rsPayments("Recorddate").Value
'                End If
'
'                paymentSQL = "INSERT INTO " & mServerD & "TblTransactionPayments (" & _
'                    "Transaction_ID, boxid, Recorddate, PointID, CurrentCashireID, PaymentID, Value, CardNo, Effect, SessionCode) VALUES ("
'                paymentSQL = paymentSQL & currentDestTransactionID & "," & _
'                    Val(rsPayments("boxid").Value & "") & "," & SQLDate(Recorddate, True) & "," & _
'                    Val(rsPayments("PointID").Value & "") & "," & Val(rsPayments("CurrentCashireID").Value & "") & "," & Val(rsPayments("PaymentID").Value & "") & "," & _
'                    Val(rsPayments("Value").Value & "") & ",'" & Replace(rsPayments("CardNo").Value & "", "'", "''") & "'," & _
'                    Val(rsPayments("Effect").Value & "") & ",'" & SessionCode & "')"
'
'                paymentsBatchSQL = paymentsBatchSQL & paymentSQL & vbCrLf
'                rsPayments.MoveNext
'            Loop
'
'            rsPayments.Close
'            Set rsPayments = Nothing
'
'            Set rsPayments = New ADODB.Recordset
'            sql = "SELECT * FROM TblSalesPayment WHERE TransID = " & Val(rsTrans("Transaction_ID").Value & "")
'            rsPayments.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'            Do While Not rsPayments.EOF
'                Dim paymentSQL2 As String
'
'                paymentSQL2 = "INSERT INTO " & mServerD & "TblSalesPayment (" & _
'                    "TransID, PaymentID, Value) VALUES (" & _
'                    currentDestTransactionID & "," & Val(rsPayments("PaymentID").Value & "") & "," & Val(rsPayments("Value").Value & "") & ")"
'
'                paymentsBatchSQL2 = paymentsBatchSQL2 & paymentSQL2 & vbCrLf
'                rsPayments.MoveNext
'            Loop
'
'            rsPayments.Close
'            Set rsPayments = Nothing
'        End If
'
'        '========================
'        ' Execute batch every 50
'        '========================
'        recCounter = recCounter + 1
'
'        If recCounter Mod BatchSize = 0 Then
'
'            BatchNo = BatchNo + 1
'            UpdateTransferCaption "Ã«—Ì  —ÕÌ· «·œ›⁄… —ﬁ„ " & BatchNo & " ≈·Ï «·”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'            WritePhaseLog "Executing batch", "BatchNo=" & BatchNo & ", RecCounter=" & recCounter & ", Total=" & TotalInvoices
'
'            If transBatchSQL <> "" Then
'                UpdateTransferCaption "Ã«—Ì  —ÕÌ· —ƒÊ” «·›Ê« Ì— ··”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'                LastSQL = transBatchSQL
'                WriteLog "Executing transBatchSQL batch", transBatchSQL
'                Cn.Execute transBatchSQL
'            End If
'
'            If detailsBatchSQL <> "" Then
'                UpdateTransferCaption "Ã«—Ì  —ÕÌ·  ›«’Ì· «·›Ê« Ì— ··”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'                LastSQL = detailsBatchSQL
'                WriteLog "Executing detailsBatchSQL batch", detailsBatchSQL
'                Cn.Execute detailsBatchSQL
'            End If
'
'            If valueAddedBatchSQL <> "" Then
'                UpdateTransferCaption "Ã«—Ì  —ÕÌ· «·÷—Ì»… Ê«·ﬁÌ„… «·„÷«›…", recCounter, TotalInvoices, SessionCode, mTimeStart
'                LastSQL = valueAddedBatchSQL
'                WriteLog "Executing valueAddedBatchSQL batch", valueAddedBatchSQL
'                Cn.Execute valueAddedBatchSQL
'            End If
'
'            If paymentsBatchSQL2 <> "" Then
'                UpdateTransferCaption "Ã«—Ì  —ÕÌ· TblSalesPayment", recCounter, TotalInvoices, SessionCode, mTimeStart
'                LastSQL = paymentsBatchSQL2
'                WriteLog "Executing paymentsBatchSQL2 batch", paymentsBatchSQL2
'                Cn.Execute paymentsBatchSQL2
'            End If
'
'            If paymentsBatchSQL <> "" Then
'                UpdateTransferCaption "Ã«—Ì  —ÕÌ· TblTransactionPayments", recCounter, TotalInvoices, SessionCode, mTimeStart
'                LastSQL = paymentsBatchSQL
'                WriteLog "Executing paymentsBatchSQL batch", paymentsBatchSQL
'                Cn.Execute paymentsBatchSQL
'            End If
'
'            transBatchSQL = ""
'            detailsBatchSQL = ""
'            valueAddedBatchSQL = ""
'            paymentsBatchSQL2 = ""
'            paymentsBatchSQL = ""
'        End If
'
'        rsTrans.MoveNext
'    Loop
'
'    '========================
'    ' Final batch
'    '========================
'    UpdateTransferCaption "Ã«—Ì  —ÕÌ· ¬Œ— œ›⁄… ≈·Ï «·”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
'    WritePhaseLog "Executing final batch", "RecCounter=" & recCounter & ", Total=" & TotalInvoices
'
'    If transBatchSQL <> "" Then
'        LastSQL = transBatchSQL
'        WriteLog "Executing transBatchSQL final", transBatchSQL
'        Cn.Execute transBatchSQL
'    End If
'
'    If detailsBatchSQL <> "" Then
'        LastSQL = detailsBatchSQL
'        WriteLog "Executing detailsBatchSQL final", detailsBatchSQL
'        Cn.Execute detailsBatchSQL
'    End If
'
'    If valueAddedBatchSQL <> "" Then
'        LastSQL = valueAddedBatchSQL
'        WriteLog "Executing valueAddedBatchSQL final", valueAddedBatchSQL
'        Cn.Execute valueAddedBatchSQL
'    End If
'
'    If paymentsBatchSQL2 <> "" Then
'        LastSQL = paymentsBatchSQL2
'        WriteLog "Executing paymentsBatchSQL2 final", paymentsBatchSQL2
'        Cn.Execute paymentsBatchSQL2
'    End If
'
'    If paymentsBatchSQL <> "" Then
'        LastSQL = paymentsBatchSQL
'        WriteLog "Executing paymentsBatchSQL final", paymentsBatchSQL
'        Cn.Execute paymentsBatchSQL
'    End If
'
'    transBatchSQL = ""
'    detailsBatchSQL = ""
'    valueAddedBatchSQL = ""
'    paymentsBatchSQL2 = ""
'    paymentsBatchSQL = ""
'
'    '========================
'    ' Reconcile and checksum
'    '========================
'    UpdateTransferCaption " „ ‰ﬁ· «·»Ì«‰« ° Ã«—Ì «·„—«Ã⁄… Ê«·„ÿ«»ﬁ…", recCounter, TotalInvoices, SessionCode, mTimeStart
'    WritePhaseLog "Start reconcile/checksum", "SessionCode=" & SessionCode
'
'    ' «·—ƒÊ”
'    Set rsCnt = POSCn.Execute( _
'      "SELECT COUNT(*) Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'")
'    SrcHeads = CLng(rsCnt!Cnt): rsCnt.Close
'
'    ' «· ›«’Ì·
'    Set rsCnt = POSCn.Execute( _
'      "SELECT COUNT(*) Cnt FROM Transaction_Details d WHERE d.Transaction_ID IN " & _
'      "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    SrcDet = CLng(rsCnt!Cnt): rsCnt.Close
'
'    ' «·ﬁÌ„… «·„÷«›…
'    Set rsCnt = POSCn.Execute( _
'      "SELECT COUNT(*) Cnt FROM TransactionValueAdded v WHERE v.Transaction_ID IN " & _
'      "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    SrcVAT = CLng(rsCnt!Cnt): rsCnt.Close
'
'    ' TransactionPayments
'    Set rsCnt = POSCn.Execute( _
'      "SELECT COUNT(*) Cnt FROM TblTransactionPayments p WHERE p.Transaction_ID IN " & _
'      "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    SrcPay = CLng(rsCnt!Cnt): rsCnt.Close
'
'    ' SalesPayment
'    Set rsCnt = POSCn.Execute( _
'      "SELECT COUNT(*) Cnt FROM TblSalesPayment s WHERE s.TransID IN " & _
'      "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    SrcPay2 = CLng(rsCnt!Cnt): rsCnt.Close
'    Set rsCnt = Nothing
'
'    UpdateTransferCaption "Ã«—Ì „—«Ã⁄… «·≈Ã„«·Ì«  Ê«·ﬂ„Ì« ", recCounter, TotalInvoices, SessionCode, mTimeStart
'
'    ' POS sums
'    Set rsCnt = POSCn.Execute( _
'        "SELECT SUM(CAST(d.Quantity AS float)) AS SumQty, " & _
'        "SUM(CAST(d.Quantity * d.Price AS decimal(18,4))) AS SumAmount " & _
'        "FROM Transaction_Details d " & _
'        "WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumQty) Then srcQty = 0# Else srcQty = CDbl(rsCnt!SumQty)
'        If IsNull(rsCnt!SumAmount) Then SrcAmount = CCur(0) Else SrcAmount = CCur(rsCnt!SumAmount)
'    End If
'    rsCnt.Close
'
'    Set rsCnt = POSCn.Execute( _
'        "SELECT SUM(CAST(v.Valu AS decimal(18,4))) AS SumVAT " & _
'        "FROM TransactionValueAdded v " & _
'        "WHERE v.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumVAT) Then SrcVATSum = CCur(0) Else SrcVATSum = CCur(rsCnt!SumVAT)
'    End If
'    rsCnt.Close
'
'    Set rsCnt = POSCn.Execute( _
'        "SELECT SUM(CAST(p.Value AS decimal(18,4))) AS SumPay " & _
'        "FROM TblTransactionPayments p " & _
'        "WHERE p.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumPay) Then SrcTPay = CCur(0) Else SrcTPay = CCur(rsCnt!SumPay)
'    End If
'    rsCnt.Close
'
'    Set rsCnt = POSCn.Execute( _
'        "SELECT SUM(CAST(s.Value AS decimal(18,4))) AS SumPay2 " & _
'        "FROM TblSalesPayment s " & _
'        "WHERE s.TransID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumPay2) Then SrcSPay = CCur(0) Else SrcSPay = CCur(rsCnt!SumPay2)
'    End If
'    rsCnt.Close
'    Set rsCnt = Nothing
'
'    ' Server sums
'    Set rsCnt = Cn.Execute( _
'        "SELECT SUM(CAST(d.Quantity AS float)) AS SumQty, " & _
'        "SUM(CAST(d.Quantity * d.Price AS decimal(18,4))) AS SumAmount " & _
'        "FROM " & mServerD & "Transaction_Details d " & _
'        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = d.Transaction_ID " & _
'        "WHERE t.SessionCode='" & SessionCode & "'")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumQty) Then dstQty = 0# Else dstQty = CDbl(rsCnt!SumQty)
'        If IsNull(rsCnt!SumAmount) Then DstAmount = CCur(0) Else DstAmount = CCur(rsCnt!SumAmount)
'    End If
'    rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'        "SELECT SUM(CAST(v.Valu AS decimal(18,4))) AS SumVAT " & _
'        "FROM " & mServerD & "TransactionValueAdded v " & _
'        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = v.Transaction_ID " & _
'        "WHERE t.SessionCode='" & SessionCode & "'")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumVAT) Then DstVATSum = CCur(0) Else DstVATSum = CCur(rsCnt!SumVAT)
'    End If
'    rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'        "SELECT SUM(CAST(p.Value AS decimal(18,4))) AS SumPay " & _
'        "FROM " & mServerD & "TblTransactionPayments p " & _
'        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = p.Transaction_ID " & _
'        "WHERE t.SessionCode='" & SessionCode & "'")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumPay) Then DstTPay = CCur(0) Else DstTPay = CCur(rsCnt!SumPay)
'    End If
'    rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'        "SELECT SUM(CAST(s.Value AS decimal(18,4))) AS SumPay2 " & _
'        "FROM " & mServerD & "TblSalesPayment s " & _
'        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = s.TransID " & _
'        "WHERE t.SessionCode='" & SessionCode & "'")
'    If Not rsCnt.EOF Then
'        If IsNull(rsCnt!SumPay2) Then DstSPay = CCur(0) Else DstSPay = CCur(rsCnt!SumPay2)
'    End If
'    rsCnt.Close
'    Set rsCnt = Nothing
'
'    epsQty = 0.0001
'    epsMoney = 0.01
'
'    If (Abs(srcQty - dstQty) > epsQty) _
'       Or (Abs(SrcAmount - DstAmount) > epsMoney) _
'       Or (Abs(SrcVATSum - DstVATSum) > epsMoney) _
'       Or (Abs(SrcTPay - DstTPay) > epsMoney) _
'       Or (Abs(SrcSPay - DstSPay) > epsMoney) Then
'
'        WriteLog "Checksum mismatch: " & _
'                 "Qty " & FormatNumber(srcQty, 6) & "/" & FormatNumber(dstQty, 6) & _
'                 ", Amount " & FormatCurrency(SrcAmount) & "/" & FormatCurrency(DstAmount) & _
'                 ", VAT " & FormatCurrency(SrcVATSum) & "/" & FormatCurrency(DstVATSum) & _
'                 ", Pay " & FormatCurrency(SrcTPay) & "/" & FormatCurrency(DstTPay) & _
'                 ", Pay2 " & FormatCurrency(SrcSPay) & "/" & FormatCurrency(DstSPay), ""
'
'        Err.Raise vbObjectError + 779, , "Checksum reconcile failed for SessionCode=" & SessionCode
'    End If
'
'    '========================
'    ' Destination counters
'    '========================
'    Set rsCnt = Cn.Execute( _
'      "SELECT COUNT(*) Cnt FROM " & mServerD & "Transactions WHERE SessionCode='" & SessionCode & "'")
'    DstHeads = CLng(rsCnt!Cnt): rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'      "SELECT COUNT(*) Cnt FROM " & mServerD & "Transaction_Details d " & _
'      "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=d.Transaction_ID " & _
'      "WHERE t.SessionCode='" & SessionCode & "'")
'    DstDet = CLng(rsCnt!Cnt): rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'      "SELECT COUNT(*) Cnt FROM " & mServerD & "TransactionValueAdded v " & _
'      "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=v.Transaction_ID " & _
'      "WHERE t.SessionCode='" & SessionCode & "'")
'    DstVAT = CLng(rsCnt!Cnt): rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'      "SELECT COUNT(*) Cnt FROM " & mServerD & "TblTransactionPayments p " & _
'      "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=p.Transaction_ID " & _
'      "WHERE t.SessionCode='" & SessionCode & "'")
'    DstPay = CLng(rsCnt!Cnt): rsCnt.Close
'
'    Set rsCnt = Cn.Execute( _
'      "SELECT COUNT(*) Cnt FROM " & mServerD & "TblSalesPayment s " & _
'      "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=s.TransID " & _
'      "WHERE t.SessionCode='" & SessionCode & "'")
'    DstPay2 = CLng(rsCnt!Cnt): rsCnt.Close
'    Set rsCnt = Nothing
'
'    UpdateTransferCaption "Ã«—Ì «· Õﬁﬁ «·‰Â«∆Ì ﬁ»· «⁄ „«œ «·‰ﬁ·", recCounter, TotalInvoices, SessionCode, mTimeStart
'
'    If (DstHeads <> SrcHeads) Or (DstDet <> SrcDet) Or (DstVAT <> SrcVAT) Or (DstPay <> SrcPay) Or (DstPay2 <> SrcPay2) Then
'        WriteLog "Reconcile mismatch: Heads " & SrcHeads & "/" & DstHeads & _
'                 ", Det " & SrcDet & "/" & DstDet & _
'                 ", VAT " & SrcVAT & "/" & DstVAT & _
'                 ", Pay " & SrcPay & "/" & DstPay & _
'                 ", Pay2 " & SrcPay2 & "/" & DstPay2, ""
'        Err.Raise vbObjectError + 778, , "Reconcile failed for SessionCode=" & SessionCode
'    End If
'
'    '========================
'    ' Final consistency check
'    '========================
'    Set rsCnt = POSCn.Execute( _
'        "SELECT COUNT(*) AS Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'")
'    srcCount = CLng(rsCnt.Fields("Cnt").Value)
'    rsCnt.Close
'    Set rsCnt = Nothing
'
'    Set rsCnt = Cn.Execute( _
'        "SELECT COUNT(*) AS Cnt FROM " & mServerD & "Transactions WHERE SessionCode='" & SessionCode & "'")
'    dstCount = CLng(rsCnt.Fields("Cnt").Value)
'    rsCnt.Close
'    Set rsCnt = Nothing
'
'    If dstCount <> srcCount Then
'        WriteLog "Session mismatch before marking Copied. src=" & srcCount & " dst=" & dstCount, ""
'        Err.Raise vbObjectError + 777, , "Mismatch between source and server counts for SessionCode=" & SessionCode
'    End If
'
'    '========================
'    ' Mark copied on source
'    '========================
'    UpdateTransferCaption " „ «· Õﬁﬁ° Ã«—Ì «⁄ „«œ «·›Ê« Ì— ﬂ„ı‰ﬁÊ·…", recCounter, TotalInvoices, SessionCode, mTimeStart
'
'    LastSQL = "UPDATE Transactions SET Copied = 1, SessionCode = NULL WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "'"
'    POSCn.Execute LastSQL
'
'    Cn.CommitTrans
'
'    '========================
'    ' Success UI
'    '========================
'    UpdateTransferCaption " „ «·‰ﬁ· »‰Ã«Õ", TotalInvoices, TotalInvoices, SessionCode, mTimeStart
'    WritePhaseLog "Transfer completed successfully", "SessionCode=" & SessionCode & ", Count=" & TotalInvoices
'    MousePointer = vbDefault
'
'    r = 1
'
'    grdInfo.TextMatrix(r, 0) = "⁄œœ «·›Ê« Ì—"
'    grdInfo.TextMatrix(r, 1) = CStr(SrcHeads)
'    grdInfo.TextMatrix(r, 2) = CStr(DstHeads)
'    grdInfo.TextMatrix(r, 3) = IIf(SrcHeads = DstHeads, "OK", "Mismatch")
'    r = r + 1
'
'    grdInfo.TextMatrix(r, 0) = "⁄œœ ”ÿÊ— «· ›«’Ì·"
'    grdInfo.TextMatrix(r, 1) = CStr(SrcDet)
'    grdInfo.TextMatrix(r, 2) = CStr(DstDet)
'    grdInfo.TextMatrix(r, 3) = IIf(SrcDet = DstDet, "OK", "Mismatch")
'    r = r + 1
'
'    grdInfo.TextMatrix(r, 0) = "«Ã„«·Ì ﬁÌ„ «·›Ê« Ì—"
'    grdInfo.TextMatrix(r, 1) = CStr(SrcAmount)
'    grdInfo.TextMatrix(r, 2) = CStr(DstAmount)
'    grdInfo.TextMatrix(r, 3) = IIf(Abs(SrcAmount - DstAmount) <= epsMoney, "OK", "ÌÊÃœ „‘ﬂ·… ›Ï «·«Ã„«·Ì« ")
'    r = r + 1
'
'    grdInfo.TextMatrix(r, 0) = "⁄œœ ”Ã·«  TransactionPayments"
'    grdInfo.TextMatrix(r, 1) = CStr(SrcPay)
'    grdInfo.TextMatrix(r, 2) = CStr(DstPay)
'    grdInfo.TextMatrix(r, 3) = IIf(SrcPay = DstPay, "OK", "Mismatch")
'    r = r + 1
'
'    grdInfo.TextMatrix(r, 0) = "⁄œœ ”Ã·«  SalesPayment"
'    grdInfo.TextMatrix(r, 1) = CStr(SrcPay2)
'    grdInfo.TextMatrix(r, 2) = CStr(DstPay2)
'    grdInfo.TextMatrix(r, 3) = IIf(SrcPay2 = DstPay2, "OK", "Mismatch")
'
'    grdInfo.ColAlignment(1) = 7
'    grdInfo.ColAlignment(2) = 7
'    grdInfo.ColAlignment(3) = 4
'
'    elapsedSec = DateDiff("s", mTimeStart, Now)
'    elapsedMin = elapsedSec \ 60
'    elapsedSec = elapsedSec Mod 60
'
'    frmPopup.ShowMessage " „ «·‰ﬁ· »‰Ã«Õ." & vbCrLf & _
'           "«·Êﬁ  «·„” €—ﬁ: " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì….", vbInformation
'
'    lblWait.Caption = " „ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄«  »‰Ã«Õ. «·Êﬁ  «·„” €—ﬁ: " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì…."
'    txtEndTime = CStr(Now)
'
'    elapsedSec = DateDiff("s", mTimeStart, Now)
'
'    Call SaveSyncLog( _
'        POSCn, Cn, _
'        SessionCode, direction, kind, _
'        mTimeStart, Now, elapsedSec, _
'        POSlServer.Text, POSDb, SysSQLServerName, ServerDb, _
'        BranchID, GetQuery, _
'        BatchSize, FetchSize, _
'        SrcHeads, DstHeads, _
'        SrcDet, DstDet, _
'        SrcVAT, DstVAT, _
'        SrcPay, DstPay, _
'        SrcPay2, DstPay2, _
'        SrcAmount, DstAmount, _
'        SrcVATSum, DstVATSum, _
'        SrcTPay, DstTPay, _
'        SrcSPay, DstSPay, _
'        True, "" _
'    )
'
'    POSname_Change
'
'EndSub:
'    On Error Resume Next
'
'    MousePointer = vbDefault
'
'    If Not rsSer Is Nothing Then
'        If rsSer.State = adStateOpen Then rsSer.Close
'        Set rsSer = Nothing
'    End If
'
'    If Not rsPayments Is Nothing Then
'        If rsPayments.State = adStateOpen Then rsPayments.Close
'        Set rsPayments = Nothing
'    End If
'
'    If Not rsValueAdded Is Nothing Then
'        If rsValueAdded.State = adStateOpen Then rsValueAdded.Close
'        Set rsValueAdded = Nothing
'    End If
'
'    If Not rsDetails Is Nothing Then
'        If rsDetails.State = adStateOpen Then rsDetails.Close
'        Set rsDetails = Nothing
'    End If
'
'    If Not rsTrans Is Nothing Then
'        If rsTrans.State = adStateOpen Then rsTrans.Close
'        Set rsTrans = Nothing
'    End If
'
'    If Not rsCnt Is Nothing Then
'        If rsCnt.State = adStateOpen Then rsCnt.Close
'        Set rsCnt = Nothing
'    End If
'
'    If Not POSCn Is Nothing Then
'        If POSCn.State = adStateOpen Then POSCn.Close
'        Set POSCn = Nothing
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    On Error Resume Next
'
'    UpdateTransferCaption "ÕœÀ Œÿ√ √À‰«¡ «·‰ﬁ· - Ã«—Ì «· —«Ã⁄", recCounter, TotalInvoices, SessionCode, mTimeStart
'
'    If Not Cn Is Nothing Then
'        If Cn.State = adStateOpen Then
'            Cn.RollbackTrans
'        End If
'    End If
'
'    If Not POSCn Is Nothing Then
'        If POSCn.State = adStateOpen Then
'            POSCn.Execute "UPDATE Transactions SET Copied = NULL, SessionCode = NULL WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "'"
'        End If
'    End If
'
'    WriteLog "ErrorHandler: " & Err.Description, LastSQL
'
'    Call SaveSyncLog( _
'        POSCn, Cn, SessionCode, direction, kind, _
'        mTimeStart, Now, DateDiff("s", mTimeStart, Now), _
'        POSlServer.Text, POSDb, SysSQLServerName, ServerDb, _
'        BranchID, GetQuery, BatchSize, FetchSize, _
'        SrcHeads, DstHeads, SrcDet, DstDet, SrcVAT, DstVAT, _
'        SrcPay, DstPay, SrcPay2, DstPay2, _
'        SrcAmount, DstAmount, SrcVATSum, DstVATSum, _
'        SrcTPay, DstTPay, SrcSPay, DstSPay, _
'        False, Err.Description)
'
'    MousePointer = vbDefault
'    lblWait.Visible = True
'    lblWait.Caption = "›‘· «·‰ﬁ·: " & Err.Description
'
'    frmPopup.ShowMessage "Œÿ√ √À‰«¡ «·‰ﬁ· —Ã«¡ «· Ê«’· „⁄ „”∆Ê·Ì «·‰Ÿ«„:" & vbCrLf & _
'                         Err.Description & vbCrLf & vbCrLf & _
'                         "¬Œ— ŒÿÊ…: " & lblWait.Caption, vbCritical
'
'    GoTo EndSub
'
'End Sub

'
'
Private Sub GetOneInvoiceSourceSummary(ByVal CnX As ADODB.Connection, ByVal SrcTransactionID As Double, _
                                       ByRef DetCnt As Long, ByRef VATCnt As Long, _
                                       ByRef TPayCnt As Long, ByRef SPayCnt As Long, _
                                       ByRef DetQty As Double, ByRef DetAmount As Currency, _
                                       ByRef VATSum As Currency, ByRef TPaySum As Currency, _
                                       ByRef SPaySum As Currency)

    DetCnt = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM Transaction_Details WHERE Transaction_ID=" & SrcTransactionID)
    VATCnt = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM TransactionValueAdded WHERE Transaction_ID=" & SrcTransactionID)
    TPayCnt = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM TblTransactionPayments WHERE Transaction_ID=" & SrcTransactionID)
    SPayCnt = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM TblSalesPayment WHERE TransID=" & SrcTransactionID)

    DetQty = ExecuteScalarDbl(CnX, "SELECT ISNULL(SUM(CAST(Quantity AS float)),0) FROM Transaction_Details WHERE Transaction_ID=" & SrcTransactionID)
    DetAmount = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Quantity * Price AS decimal(18,4))),0) FROM Transaction_Details WHERE Transaction_ID=" & SrcTransactionID)
    VATSum = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Valu AS decimal(18,4))),0) FROM TransactionValueAdded WHERE Transaction_ID=" & SrcTransactionID)
    TPaySum = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Value AS decimal(18,4))),0) FROM TblTransactionPayments WHERE Transaction_ID=" & SrcTransactionID)
    SPaySum = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Value AS decimal(18,4))),0) FROM TblSalesPayment WHERE TransID=" & SrcTransactionID)

End Sub

Private Sub DebugLogDestOneInvoiceSummary(ByVal CnX As ADODB.Connection, ByVal DestTransactionID As String, _
                                          ByVal DebugLogFile As String, ByVal Title As String)

    Dim s As String
    Dim nDet As Long
    Dim nVAT As Long
    Dim nTPay As Long
    Dim nSPay As Long
    Dim q As Double
    Dim a As Currency
    Dim v As Currency
    Dim p1 As Currency
    Dim p2 As Currency

    If Trim$(DestTransactionID) = "" Then Exit Sub

    nDet = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM dbo.Transaction_Details WHERE Transaction_ID=" & DestTransactionID)
    nVAT = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM dbo.TransactionValueAdded WHERE Transaction_ID=" & DestTransactionID)
    nTPay = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM dbo.TblTransactionPayments WHERE Transaction_ID=" & DestTransactionID)
    nSPay = ExecuteScalarLng(CnX, "SELECT COUNT(*) FROM dbo.TblSalesPayment WHERE TransID=" & DestTransactionID)

    q = ExecuteScalarDbl(CnX, "SELECT ISNULL(SUM(CAST(Quantity AS float)),0) FROM dbo.Transaction_Details WHERE Transaction_ID=" & DestTransactionID)
    a = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Quantity * Price AS decimal(18,4))),0) FROM dbo.Transaction_Details WHERE Transaction_ID=" & DestTransactionID)
    v = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Valu AS decimal(18,4))),0) FROM dbo.TransactionValueAdded WHERE Transaction_ID=" & DestTransactionID)
    p1 = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Value AS decimal(18,4))),0) FROM dbo.TblTransactionPayments WHERE Transaction_ID=" & DestTransactionID)
    p2 = ExecuteScalarCur(CnX, "SELECT ISNULL(SUM(CAST(Value AS decimal(18,4))),0) FROM dbo.TblSalesPayment WHERE TransID=" & DestTransactionID)

    s = Title & " => DestID=" & DestTransactionID & _
        " | DetCnt=" & nDet & _
        " | VATCnt=" & nVAT & _
        " | TPayCnt=" & nTPay & _
        " | SPayCnt=" & nSPay & _
        " | Qty=" & CStr(q) & _
        " | Amount=" & CStr(a) & _
        " | VATSum=" & CStr(v) & _
        " | TPaySum=" & CStr(p1) & _
        " | SPaySum=" & CStr(p2)

    DebugWriteLine DebugLogFile, s

End Sub

Private Function ExecuteScalarLng(ByVal CnX As ADODB.Connection, ByVal SQLText As String) As Long
    Dim rs As ADODB.Recordset

    On Error GoTo EH
    Set rs = CnX.Execute(SQLText)

    If rs.EOF Then
        ExecuteScalarLng = 0
    ElseIf IsNull(rs.Fields(0).Value) Then
        ExecuteScalarLng = 0
    Else
        ExecuteScalarLng = CLng(rs.Fields(0).Value)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    ExecuteScalarLng = 0
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
    End If
    Set rs = Nothing
End Function
'
Private Sub Command14_Click()

    On Error GoTo ErrorHandler

    '========================
    ' Declarations
    '========================
    Dim POSCn As ADODB.Connection
    Dim rsCnt As ADODB.Recordset
    Dim rsTrans As ADODB.Recordset
    Dim rsDetails As ADODB.Recordset
    Dim rsValueAdded As ADODB.Recordset
    Dim rsPayments As ADODB.Recordset
    Dim rsSer As ADODB.Recordset

    Dim BatchSize As Integer
    Dim recCounter As Long
    Dim recCount As Long
    Dim BatchNo As Long
    Dim TotalInvoices As Long

    Dim LastSQL As String
    Dim sql As String
    Dim transSQL As String
    Dim detailSQL As String
    Dim valueSQL As String
    Dim paymentSQL As String
    Dim paymentSQL2 As String

    Dim SessionCode As String
    Dim CurrentInvoiceNo As String
    Dim CurrentStage As String
    Dim CurrentSrcTransactionID As Double
    Dim currentDestTransactionID As String
    Dim ReconcileMsg As String

    Dim mTimeStart As Date
    Dim direction As String
    Dim kind As String
    Dim FetchSize As Long

    Dim CarOilChangeDate As Date
    Dim RecTime As Date
    Dim mTimeIn As String
    Dim ActualDeliveryDate As Date
    Dim LatestDeliveryDate As Date
    Dim FromTransaction_Date As Date
    Dim FromTransaction_Type As Long
    Dim FromTransaction_ID As Double

    Dim PayMentType As Long
    Dim cusID As Long
    Dim BranchID As Integer
    Dim BoxID As Long
    Dim BillBasedOn As Double
    Dim VAT As Double
    Dim VATYou As Double
    Dim NoteId As Long
    Dim Trans_DiscountType As Long
    Dim Trans_Discount As Double
    Dim TaxValue As Double
    Dim order_no As Long
    Dim SaleType As Long
    Dim TaxAddValue As Double
    Dim NetValue As Double
    Dim Transaction_NetValue As Double
    Dim DepandToConv As Long
    Dim CarTypeID As Long
    Dim OilsTypesID As Long
    Dim YearFact As Long
    Dim FixesAssetsID As Long
    Dim ColorID2 As Long
    Dim KM As Double
    Dim PPointID As Long
    Dim SupplerID As Long
    Dim Ser As Long
    Dim CarCurrentValue As Double
    Dim CarPrevValue As Double
    Dim CarEnginoil As Double
    Dim CarGearOil As Double
    Dim InvoiceTypeCodeID As Long
    Dim storeID As Variant
    Dim userID As Variant
    Dim Emp_ID As Variant
    Dim NoteSerial As String
    Dim NoteSerial1 As String
    Dim TransactionComment As String
    Dim CashCustomerName As String
    Dim CashCustomerPhone As String
    Dim cleanCashCustomerName As String
    Dim PlateNo As String
    Dim Shaseh As String
    Dim CarMeter As String
    Dim Chasee As String
    Dim Phone2 As String
    Dim CIBAN As String
    Dim InvoiceTypeCodename As String
    Dim DocumentCurrencyCode As String
    Dim TaxCurrencyCode As String
    Dim paymentnote As String
    Dim PaymentMeansCode As String

    ' Counters for reconcile
    Dim SrcHeads As Long
    Dim SrcDet As Long
    Dim SrcVAT As Long
    Dim SrcPay As Long
    Dim SrcPay2 As Long
    Dim DstHeads As Long
    Dim DstDet As Long
    Dim DstVAT As Long
    Dim DstPay As Long
    Dim DstPay2 As Long
    Dim srcCount As Long
    Dim dstCount As Long

    ' Checksums
    Dim srcQty As Double
    Dim dstQty As Double
    Dim SrcAmount As Currency
    Dim DstAmount As Currency
    Dim SrcVATSum As Currency
    Dim DstVATSum As Currency
    Dim SrcTPay As Currency
    Dim DstTPay As Currency
    Dim SrcSPay As Currency
    Dim DstSPay As Currency
    Dim epsQty As Double
    Dim epsMoney As Currency

    ' Per invoice summary
    Dim OneSrcDetCnt As Long
    Dim OneSrcVATCnt As Long
    Dim OneSrcPayCnt As Long
    Dim OneSrcPay2Cnt As Long
    Dim OneSrcQty As Double
    Dim OneSrcAmount As Currency
    Dim OneSrcVAT As Currency
    Dim OneSrcPay As Currency
    Dim OneSrcPay2 As Currency

    Dim OneDstDetCnt As Long
    Dim OneDstVATCnt As Long
    Dim OneDstPayCnt As Long
    Dim OneDstPay2Cnt As Long
    Dim OneDstQty As Double
    Dim OneDstAmount As Currency
    Dim OneDstVAT As Currency
    Dim OneDstPay As Currency
    Dim OneDstPay2 As Currency

    ' Misc
    Dim tmpRecTime As Variant
    Dim v As Variant
    Dim r As Integer
    Dim mServerD As String
    Dim inTx As Boolean
    Dim Recorddate As Date
    Dim POSBillTypeVal As Long
    Dim MsgErr As String

    '========================
    ' Initial values
    '========================
    BatchSize = 50
    recCounter = 0
    recCount = 0
    BatchNo = 0
    FetchSize = 0

    LastSQL = ""
    direction = "POS->Server"
    kind = "Sales"
    mServerD = "dbo."
    inTx = False
    ReconcileMsg = ""
    CurrentStage = "Init"
    CurrentInvoiceNo = ""
    CurrentSrcTransactionID = 0
    currentDestTransactionID = ""

    mTimeStart = Now
    txtStartTime = mTimeStart

    '========================
    ' Validation
    '========================
    If Trim$(POSlServer.Text) = "" Then
        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
        Exit Sub
    End If

    If ConnectionFirst = False Then Exit Sub

    lblWait.Visible = True
    lblWait.Caption = "Ã«—Ì »œ¡ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄« ..."
    DoEvents
    MousePointer = vbHourglass

    SessionCode = Format(Now, "yyyymmddhhnnss")

    WritePhaseLog "Start Command14_Click", "SessionCode=" & SessionCode
    WriteLog "DEBUG Connection", _
             "ServerDb=" & ServerDb & _
             " | POSDb=" & POSDb & _
             " | SysSQLServerName=" & SysSQLServerName & _
             " | POSServer=" & POSlServer.Text & _
             " | GetQuery=" & GetQuery

    UpdateTransferCaption "Ã«—Ì  ÃÂÌ“ Ã·”… «·‰ﬁ·", 0, 0, SessionCode, mTimeStart

    '========================
    ' Open local POS connection
    '========================
    CurrentStage = "Open POS connection"

    Set POSCn = New ADODB.Connection
    POSCn.CursorLocation = adUseServer
    POSCn.ConnectionTimeout = 5000
    POSCn.CommandTimeout = 5000
    POSCn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
                             ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                             ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer.Text
    POSCn.Open

    WriteLog "DEBUG POSCn.ConnectionString", POSCn.ConnectionString
    If Not Cn Is Nothing Then
        WriteLog "DEBUG Cn.ConnectionString", Cn.ConnectionString
    End If

    Cn.CursorLocation = adUseServer

    '========================
    ' Tag source transactions
    '========================
    CurrentStage = "Tag source transactions"
    LastSQL = "UPDATE Transactions SET SessionCode = '" & SessionCode & "' WHERE IsNull(Copied,0)=0 AND " & GetQuery
    POSCn.Execute LastSQL

    WritePhaseLog "Tagging source transactions", "SessionCode=" & SessionCode

    '========================
    ' Get total count first
    '========================
    CurrentStage = "Count source transactions"
    LastSQL = "SELECT COUNT(*) AS Cnt FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND " & GetQuery

    Set rsCnt = POSCn.Execute(LastSQL)
    TotalInvoices = CLng(Val(rsCnt!Cnt & ""))
    rsCnt.Close
    Set rsCnt = Nothing

    WritePhaseLog "Invoices selected", "Count=" & TotalInvoices
    UpdateTransferCaption " „ «·⁄ÀÊ— ⁄·Ï ›Ê« Ì— ··‰ﬁ·", 0, TotalInvoices, SessionCode, mTimeStart

    If TotalInvoices = 0 Then
        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
        GoTo EndSub
    End If

    '========================
    ' Open source transactions
    '========================
    CurrentStage = "Open source transactions"
    LastSQL = "SELECT * FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND " & GetQuery & " ORDER BY Transaction_ID"

    Set rsTrans = New ADODB.Recordset
    rsTrans.Open LastSQL, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rsTrans.EOF Then
        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
        GoTo EndSub
    End If

    '========================
    ' Start SQL transaction on central server
    '========================
    CurrentStage = "Begin server transaction"
    LastSQL = "SET XACT_ABORT ON;"
    Cn.Execute LastSQL

    Cn.BeginTrans
    inTx = True

    '========================
    ' Main Loop
    '========================
    Do While Not rsTrans.EOF

        CurrentSrcTransactionID = Val(rsTrans("Transaction_ID").Value & "")
        currentDestTransactionID = ""

        If Trim$(rsTrans("NoteSerial1").Value & "") <> "" Then
            CurrentInvoiceNo = Trim$(rsTrans("NoteSerial1").Value & "")
        Else
            CurrentInvoiceNo = CStr(CurrentSrcTransactionID)
        End If

        CurrentStage = "Read header values"

        If recCounter = 0 Or recCounter Mod 5 = 0 Then
            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·›« Ê—… —ﬁ„ " & CurrentInvoiceNo, recCounter + 1, TotalInvoices, SessionCode, mTimeStart
        End If

        recCount = recCount + 1
        recCounter = recCounter + 1

        PayMentType = Val(rsTrans("PaymentType").Value & "")
        FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
        cusID = Val(rsTrans("CusID").Value & "")
        storeID = Val(rsTrans("StoreID").Value & "")
        userID = Val(rsTrans("UserID").Value & "")
        Emp_ID = Val(rsTrans("Emp_ID").Value & "")
        BranchID = Val(rsTrans("BranchID").Value & "")
        BoxID = Val(rsTrans("BoxID").Value & "")
        BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
        mTimeIn = Trim$(rsTrans("TimeIn").Value & "")
        VAT = Val(rsTrans("VAT").Value & "")
        VATYou = Val(rsTrans("VATYou").Value & "")
        NoteSerial = rsTrans("NoteSerial").Value & ""
        NoteSerial1 = rsTrans("NoteSerial1").Value & ""
        NoteId = Val(rsTrans("NoteId").Value & "")
        TransactionComment = rsTrans("TransactionComment").Value & ""
        Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
        FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
        Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
        TaxValue = Val(rsTrans("TaxValue").Value & "")
        order_no = Val(rsTrans("order_no").Value & "")
        SaleType = Val(rsTrans("SaleType").Value & "")
        CashCustomerName = rsTrans("CashCustomerName").Value & ""
        cleanCashCustomerName = CashCustomerName
        TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
        CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
        NetValue = Val(rsTrans("NetValue").Value & "")
        Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
        DepandToConv = Val(rsTrans("DepandToConv").Value & "")
        CarTypeID = Val(rsTrans("CarTypeID").Value & "")
        PlateNo = rsTrans("PlateNo").Value & ""
        OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
        YearFact = Val(rsTrans("YearFact").Value & "")
        Shaseh = rsTrans("Shaseh").Value & ""
        CarMeter = rsTrans("CarMeter").Value & ""
        FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
        ColorID2 = Val(rsTrans("ColorID2").Value & "")
        KM = Val(rsTrans("KM").Value & "")
        Chasee = rsTrans("Chasee").Value & ""
        PPointID = Val(rsTrans("PPointID").Value & "")
        Phone2 = rsTrans("Phone2").Value & ""
        SupplerID = Val(rsTrans("SupplerID").Value & "")
        Ser = Val(rsTrans("Ser").Value & "")
        CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
        CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
        CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
        CarGearOil = Val(rsTrans("CarGearOil").Value & "")
        CIBAN = rsTrans("CIBAN").Value & ""
        InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
        InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
        DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
        TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
        paymentnote = rsTrans("paymentnote").Value & ""
        PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""
        POSBillTypeVal = Val(rsTrans("POSBillType").Value & "")

        FromTransaction_Date = rsTrans("Transaction_Date").Value

        If Trim$(rsTrans("CarOilChangeDate").Value & "") = "" Then
            CarOilChangeDate = Date
        Else
            CarOilChangeDate = rsTrans("CarOilChangeDate").Value
        End If

        tmpRecTime = rsTrans("RecTime").Value
        v = tmpRecTime
        If IsDate(v) Then
            If Year(CDate(v)) = 1899 And Month(CDate(v)) = 12 And Day(CDate(v)) = 30 Then
                RecTime = Time
            Else
                RecTime = CDate(v)
            End If
        Else
            RecTime = Time
        End If

        If Trim$(rsTrans("ActualDeliveryDate").Value & "") = "" Then
            ActualDeliveryDate = Date
        Else
            ActualDeliveryDate = rsTrans("ActualDeliveryDate").Value
        End If

        If Trim$(rsTrans("LatestDeliveryDate").Value & "") = "" Then
            LatestDeliveryDate = ActualDeliveryDate
        Else
            LatestDeliveryDate = rsTrans("LatestDeliveryDate").Value
        End If

        If POSBillTypeVal = 0 Then
            NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
            NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
        End If

        TransactionComment = " ›« Ê—… „‰ﬁÊ·… „‰ " & POSname.Text & "    —ﬁ„ «·›« Ê—… " & NoteSerial1

        '========================
        ' Log source invoice summary
        '========================
        GetOneInvoiceSummary POSCn, CurrentSrcTransactionID, False, _
                             OneSrcDetCnt, OneSrcVATCnt, OneSrcPayCnt, OneSrcPay2Cnt, _
                             OneSrcQty, OneSrcAmount, OneSrcVAT, OneSrcPay, OneSrcPay2

        WriteLog "Invoice Start", _
                 "SrcID=" & CurrentSrcTransactionID & _
                 " | InvoiceNo=" & CurrentInvoiceNo & _
                 " | Type=" & FromTransaction_Type & _
                 " | SrcDet=" & OneSrcDetCnt & _
                 " | SrcVAT=" & OneSrcVATCnt & _
                 " | SrcTPay=" & OneSrcPayCnt & _
                 " | SrcSPay=" & OneSrcPay2Cnt & _
                 " | SrcQty=" & CStr(OneSrcQty) & _
                 " | SrcAmount=" & CStr(OneSrcAmount)

        If (FromTransaction_Type <> 21 And FromTransaction_Type <> 9) And (OneSrcPayCnt > 0 Or OneSrcPay2Cnt > 0) Then
            WriteLog "WARNING", "Source invoice has payments but type is not 21/9. SrcID=" & CurrentSrcTransactionID & " | Type=" & FromTransaction_Type
        End If

        '========================
        ' Reserve destination Transaction_ID
        '========================
        CurrentStage = "Reserve destination Transaction_ID"
        LastSQL = "EXEC dbo.ReserveTransactionId"

'        Set rsSer = Cn.Execute(LastSQL)
        
        Dim cmdSer As ADODB.Command
Set cmdSer = New ADODB.Command

Set cmdSer.ActiveConnection = Cn
cmdSer.CommandType = adCmdText
cmdSer.CommandText = "EXEC dbo.ReserveTransactionId"

Set rsSer = cmdSer.Execute

        If rsSer.EOF Then
            Err.Raise vbObjectError + 500, , "·„ Ì „ ≈—Ã«⁄ Transaction_ID ÃœÌœ „‰ «·”Ì—›—"
        End If

        currentDestTransactionID = rsSer.Fields("NewId").Value & ""

        rsSer.Close
        Set rsSer = Nothing

        WriteLog "Reserved destination ID", "SrcID=" & CurrentSrcTransactionID & " | DestID=" & currentDestTransactionID

        '========================
        ' Insert Transactions Header
        '========================
        CurrentStage = "Insert Transactions Header"

        transSQL = "INSERT INTO " & mServerD & "Transactions (" & _
                   "Transaction_ID, Transaction_Date, TimeIn, TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
                   "CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, NoteSerial1, " & _
                   "NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, OldNoteserial, OldNoteId, " & _
                   "OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, SaleType, CashCustomerName, " & _
                   "TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, DepandToConv, CarTypeID, " & _
                   "PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, ColorID2, KM, Chasee, PPointID, Phone2, " & _
                   "SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, CarGearOil, CarOilChangeDate, CIBAN, RecTime, " & _
                   "ActualDeliveryDate, LatestDeliveryDate, InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, " & _
                   "TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("

        transSQL = transSQL & currentDestTransactionID & "," & _
                   SQLDate(FromTransaction_Date, True) & ",'" & Replace(Trim$(mTimeIn), "'", "''") & "'," & _
                   Val(rsTrans("TypeInvoice").Value & "") & ",'" & Replace(rsTrans("Transaction_Serial").Value & "", "'", "''") & "'," & _
                   FromTransaction_Type & "," & PayMentType & "," & cusID & "," & storeID & "," & userID & "," & _
                   Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & ",'" & _
                   Replace(NoteSerial, "'", "''") & "','" & Replace(NoteSerial1, "'", "''") & "'," & NoteId & ",1,'" & _
                   Replace(TransactionComment, "'", "''") & "','" & SessionCode & "'," & _
                   IIf(POSBillTypeVal = 0, 1, POSBillTypeVal) & ",'" & _
                   Replace(rsTrans("NoteSerial1").Value & "", "'", "''") & "','" & Replace(Trim$(rsTrans("NoteSerial").Value & ""), "'", "''") & "'," & _
                   Val(rsTrans("NoteId").Value & "") & "," & Val(rsTrans("Transaction_ID").Value & "") & "," & _
                   Trans_DiscountType & "," & Trans_Discount & "," & TaxValue & ",'" & order_no & "'," & _
                   SaleType & ",'" & Replace(cleanCashCustomerName, "'", "''") & "'," & TaxAddValue & ",'" & _
                   Replace(CashCustomerPhone, "'", "''") & "'," & SQLDate(rsTrans("last_changed").Value, True) & ","

        transSQL = transSQL & NetValue & "," & Transaction_NetValue & "," & _
                   IIf(DepandToConv <> 0, 1, 0) & "," & CarTypeID & ",'" & Replace(PlateNo, "'", "''") & "'," & _
                   OilsTypesID & "," & YearFact & ",'" & Replace(Shaseh, "'", "''") & "','" & Replace(CarMeter, "'", "''") & "'," & _
                   FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Replace(Chasee, "'", "''") & "'," & PPointID & ",'" & _
                   Replace(Phone2, "'", "''") & "'," & SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & _
                   CarEnginoil & "," & CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & Replace(CIBAN, "'", "''") & "'," & _
                   SQLDate(RecTime, True) & "," & SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
                   InvoiceTypeCodeID & ",'" & Replace(InvoiceTypeCodename, "'", "''") & "','" & Replace(DocumentCurrencyCode, "'", "''") & "','" & _
                   Replace(TaxCurrencyCode, "'", "''") & "','" & Replace(paymentnote, "'", "''") & "','" & Replace(PaymentMeansCode, "'", "''") & "')"

        LastSQL = transSQL
        Cn.Execute transSQL

        GetOneInvoiceSummary Cn, CDbl(currentDestTransactionID), True, _
                             OneDstDetCnt, OneDstVATCnt, OneDstPayCnt, OneDstPay2Cnt, _
                             OneDstQty, OneDstAmount, OneDstVAT, OneDstPay, OneDstPay2

        WriteLog "After Header", _
                 "SrcID=" & CurrentSrcTransactionID & _
                 " | DestID=" & currentDestTransactionID & _
                 " | DstDet=" & OneDstDetCnt & _
                 " | DstVAT=" & OneDstVATCnt & _
                 " | DstTPay=" & OneDstPayCnt & _
                 " | DstSPay=" & OneDstPay2Cnt

        '========================
        ' Insert details
        '========================
        CurrentStage = "Insert Transaction_Details"

        If recCounter = 1 Or recCounter Mod 10 = 0 Then
            UpdateTransferCaption "Ã«—Ì ﬁ—«¡…  ›«’Ì· «·›Ê« Ì—", recCounter, TotalInvoices, SessionCode, mTimeStart
        End If

        sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & CurrentSrcTransactionID
        Set rsDetails = New ADODB.Recordset
        rsDetails.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

        Do While Not rsDetails.EOF

            detailSQL = "INSERT INTO " & mServerD & "Transaction_Details (" & _
                        "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
                        "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
                        "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES (" & _
                        currentDestTransactionID & "," & Val(rsDetails("Item_ID").Value & "") & "," & Val(rsDetails("ItemCase").Value & "") & "," & _
                        Val(rsDetails("Quantity").Value & "") & "," & Val(rsDetails("Price").Value & "") & "," & Val(rsDetails("ItemDiscountType").Value & "") & "," & _
                        Val(rsDetails("ItemDiscount").Value & "") & "," & Val(rsDetails("ShowQty").Value & "") & "," & Val(rsDetails("showPrice").Value & "") & "," & _
                        Val(rsDetails("UnitId").Value & "") & "," & Val(rsDetails("ColorID").Value & "") & "," & Val(rsDetails("ItemSize").Value & "") & "," & _
                        Val(rsDetails("ClassId").Value & "") & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo").Value & "") & "," & _
                        Val(rsDetails("PumpId").Value & "") & "," & Val(rsDetails("PrevQty").Value & "") & ",'" & Replace(Trim$(rsDetails("PrintName").Value & ""), "'", "''") & "'," & _
                        Val(rsDetails("Cash").Value & "") & "," & Val(rsDetails("Mada").Value & "") & "," & Val(rsDetails("Visa").Value & "") & "," & _
                        Val(rsDetails("Deferred").Value & "") & "," & Val(rsDetails("AmountH").Value & "") & "," & Val(rsDetails("AmountHComm").Value & "")

            detailSQL = detailSQL & ",'" & Replace(Trim$(rsDetails("DetailsPump").Value & ""), "'", "''") & "','" & _
                        Replace(Trim$(rsDetails("Account_CodeComm").Value & ""), "'", "''") & "','" & _
                        Replace(Trim$(rsDetails("Account_Code").Value & ""), "'", "''") & "'," & IIf(Val(rsDetails("IsOther").Value & "") <> 0, 1, 0) & ")"

            LastSQL = detailSQL
            Cn.Execute detailSQL

            rsDetails.MoveNext
        Loop

        rsDetails.Close
        Set rsDetails = Nothing

        GetOneInvoiceSummary Cn, CDbl(currentDestTransactionID), True, _
                             OneDstDetCnt, OneDstVATCnt, OneDstPayCnt, OneDstPay2Cnt, _
                             OneDstQty, OneDstAmount, OneDstVAT, OneDstPay, OneDstPay2

        WriteLog "After Details", _
                 "SrcID=" & CurrentSrcTransactionID & _
                 " | DestID=" & currentDestTransactionID & _
                 " | DstDet=" & OneDstDetCnt & _
                 " | DstQty=" & CStr(OneDstQty) & _
                 " | DstAmount=" & CStr(OneDstAmount)

        '========================
        ' Insert VAT
        '========================
        CurrentStage = "Insert TransactionValueAdded"

        If recCounter = 1 Or recCounter Mod 10 = 0 Then
            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·÷—Ì»… Ê«·ﬁÌ„… «·„÷«›…", recCounter, TotalInvoices, SessionCode, mTimeStart
        End If

        sql = "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & CurrentSrcTransactionID
        Set rsValueAdded = New ADODB.Recordset
        rsValueAdded.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

        Do While Not rsValueAdded.EOF

            valueSQL = "INSERT INTO " & mServerD & "TransactionValueAdded (" & _
                       "Transaction_ID, ItemID, Vatyo, VAT, Valu, selectd, Transaction_Type, SessionCode) VALUES (" & _
                       currentDestTransactionID & "," & Val(rsValueAdded("ItemID").Value & "") & "," & _
                       Val(rsValueAdded("Vatyo").Value & "") & "," & Val(rsValueAdded("Vat").Value & "") & "," & _
                       Val(rsValueAdded("Valu").Value & "") & "," & Val(rsValueAdded("selectd").Value & "") & "," & _
                       Val(rsValueAdded("Transaction_Type").Value & "") & ",'" & SessionCode & "')"

            LastSQL = valueSQL
            Cn.Execute valueSQL

            rsValueAdded.MoveNext
        Loop

        rsValueAdded.Close
        Set rsValueAdded = Nothing

        GetOneInvoiceSummary Cn, CDbl(currentDestTransactionID), True, _
                             OneDstDetCnt, OneDstVATCnt, OneDstPayCnt, OneDstPay2Cnt, _
                             OneDstQty, OneDstAmount, OneDstVAT, OneDstPay, OneDstPay2

        WriteLog "After VAT", _
                 "SrcID=" & CurrentSrcTransactionID & _
                 " | DestID=" & currentDestTransactionID & _
                 " | DstVATCnt=" & OneDstVATCnt & _
                 " | DstVAT=" & CStr(OneDstVAT)

        '========================
        ' Insert payments
        '========================
        CurrentStage = "Insert Payments"

        If recCounter = 1 Or recCounter Mod 10 = 0 Then
            UpdateTransferCaption "Ã«—Ì ﬁ—«¡… «·„œ›Ê⁄« ", recCounter, TotalInvoices, SessionCode, mTimeStart
        End If

        If FromTransaction_Type = 21 Or FromTransaction_Type = 9 Then

            sql = "SELECT * FROM TblTransactionPayments WHERE Transaction_ID = " & CurrentSrcTransactionID
            Set rsPayments = New ADODB.Recordset
            rsPayments.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

            Do While Not rsPayments.EOF

                If IsNull(rsPayments("Recorddate").Value) Or Trim$(rsPayments("Recorddate").Value & "") = "" Then
                    Recorddate = Now
                Else
                    Recorddate = rsPayments("Recorddate").Value
                End If

                paymentSQL = "INSERT INTO " & mServerD & "TblTransactionPayments (" & _
                             "Transaction_ID, boxid, Recorddate, PointID, CurrentCashireID, PaymentID, Value, CardNo, Effect, SessionCode) VALUES (" & _
                             currentDestTransactionID & "," & Val(rsPayments("boxid").Value & "") & "," & SQLDate(Recorddate, True) & "," & _
                             Val(rsPayments("PointID").Value & "") & "," & Val(rsPayments("CurrentCashireID").Value & "") & "," & _
                             Val(rsPayments("PaymentID").Value & "") & "," & Val(rsPayments("Value").Value & "") & ",'" & _
                             Replace(rsPayments("CardNo").Value & "", "'", "''") & "'," & Val(rsPayments("Effect").Value & "") & ",'" & SessionCode & "')"

                LastSQL = paymentSQL
                Cn.Execute paymentSQL

                rsPayments.MoveNext
            Loop

            rsPayments.Close
            Set rsPayments = Nothing

            sql = "SELECT * FROM TblSalesPayment WHERE TransID = " & CurrentSrcTransactionID
            Set rsPayments = New ADODB.Recordset
            rsPayments.Open sql, POSCn, adOpenForwardOnly, adLockReadOnly, adCmdText

            Do While Not rsPayments.EOF

                paymentSQL2 = "INSERT INTO " & mServerD & "TblSalesPayment (" & _
                              "TransID, PaymentID, Value) VALUES (" & _
                              currentDestTransactionID & "," & Val(rsPayments("PaymentID").Value & "") & "," & Val(rsPayments("Value").Value & "") & ")"

                LastSQL = paymentSQL2
                Cn.Execute paymentSQL2

                rsPayments.MoveNext
            Loop

            rsPayments.Close
            Set rsPayments = Nothing
        Else
            WriteLog "Payments Skipped", "SrcID=" & CurrentSrcTransactionID & " | Type=" & FromTransaction_Type
        End If

        GetOneInvoiceSummary Cn, CDbl(currentDestTransactionID), True, _
                             OneDstDetCnt, OneDstVATCnt, OneDstPayCnt, OneDstPay2Cnt, _
                             OneDstQty, OneDstAmount, OneDstVAT, OneDstPay, OneDstPay2

        WriteLog "After Payments", _
                 "SrcID=" & CurrentSrcTransactionID & _
                 " | DestID=" & currentDestTransactionID & _
                 " | DstTPayCnt=" & OneDstPayCnt & _
                 " | DstSPayCnt=" & OneDstPay2Cnt & _
                 " | DstTPay=" & CStr(OneDstPay) & _
                 " | DstSPay=" & CStr(OneDstPay2)

        If recCounter Mod BatchSize = 0 Then
            BatchNo = BatchNo + 1
            UpdateTransferCaption " „  —ÕÌ· «·œ›⁄… —ﬁ„ " & BatchNo & " ≈·Ï «·”Ì—›—", recCounter, TotalInvoices, SessionCode, mTimeStart
            WritePhaseLog "Executed batch", "BatchNo=" & BatchNo & ", RecCounter=" & recCounter & ", Total=" & TotalInvoices
        End If

        rsTrans.MoveNext
    Loop

    '========================
    ' Reconcile and checksum
    '========================
    CurrentStage = "Reconcile/checksum"
    UpdateTransferCaption " „ ‰ﬁ· «·»Ì«‰« ° Ã«—Ì «·„—«Ã⁄… Ê«·„ÿ«»ﬁ…", recCounter, TotalInvoices, SessionCode, mTimeStart
    WritePhaseLog "Start reconcile/checksum", "SessionCode=" & SessionCode

    ' Source counters
    SrcHeads = ExecuteScalarLng(POSCn, _
        "SELECT COUNT(*) FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'")

    SrcDet = ExecuteScalarLng(POSCn, _
        "SELECT COUNT(*) FROM Transaction_Details d WHERE d.Transaction_ID IN " & _
        "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")

    SrcVAT = ExecuteScalarLng(POSCn, _
        "SELECT COUNT(*) FROM TransactionValueAdded v WHERE v.Transaction_ID IN " & _
        "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")

    ' IMPORTANT:
    ' payments are inserted only for Transaction_Type IN (21,9)
    SrcPay = ExecuteScalarLng(POSCn, _
        "SELECT COUNT(*) FROM TblTransactionPayments p WHERE p.Transaction_ID IN " & _
        "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND Transaction_Type IN (21,9))")

    SrcPay2 = ExecuteScalarLng(POSCn, _
        "SELECT COUNT(*) FROM TblSalesPayment s WHERE s.TransID IN " & _
        "(SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND Transaction_Type IN (21,9))")

    ' Source sums
    srcQty = ExecuteScalarDbl(POSCn, _
        "SELECT ISNULL(SUM(CAST(d.Quantity AS float)),0) FROM Transaction_Details d " & _
        "WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")

    SrcAmount = ExecuteScalarCur(POSCn, _
        "SELECT ISNULL(SUM(CAST(d.Quantity * d.Price AS decimal(18,4))),0) FROM Transaction_Details d " & _
        "WHERE d.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")

    SrcVATSum = ExecuteScalarCur(POSCn, _
        "SELECT ISNULL(SUM(CAST(v.Valu AS decimal(18,4))),0) FROM TransactionValueAdded v " & _
        "WHERE v.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "')")

    SrcTPay = ExecuteScalarCur(POSCn, _
        "SELECT ISNULL(SUM(CAST(p.Value AS decimal(18,4))),0) FROM TblTransactionPayments p " & _
        "WHERE p.Transaction_ID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND Transaction_Type IN (21,9))")

    SrcSPay = ExecuteScalarCur(POSCn, _
        "SELECT ISNULL(SUM(CAST(s.Value AS decimal(18,4))),0) FROM TblSalesPayment s " & _
        "WHERE s.TransID IN (SELECT Transaction_ID FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "' AND Transaction_Type IN (21,9))")

    ' Destination counters
    DstHeads = ExecuteScalarLng(Cn, _
        "SELECT COUNT(*) FROM " & mServerD & "Transactions WHERE SessionCode='" & SessionCode & "'")

    DstDet = ExecuteScalarLng(Cn, _
        "SELECT COUNT(*) FROM " & mServerD & "Transaction_Details d " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=d.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "'")

    DstVAT = ExecuteScalarLng(Cn, _
        "SELECT COUNT(*) FROM " & mServerD & "TransactionValueAdded v " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=v.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "'")

    DstPay = ExecuteScalarLng(Cn, _
        "SELECT COUNT(*) FROM " & mServerD & "TblTransactionPayments p " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=p.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "' AND t.Transaction_Type IN (21,9)")

    DstPay2 = ExecuteScalarLng(Cn, _
        "SELECT COUNT(*) FROM " & mServerD & "TblSalesPayment s " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID=s.TransID " & _
        "WHERE t.SessionCode='" & SessionCode & "' AND t.Transaction_Type IN (21,9)")

    ' Destination sums
    dstQty = ExecuteScalarDbl(Cn, _
        "SELECT ISNULL(SUM(CAST(d.Quantity AS float)),0) FROM " & mServerD & "Transaction_Details d " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = d.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "'")

    DstAmount = ExecuteScalarCur(Cn, _
        "SELECT ISNULL(SUM(CAST(d.Quantity * d.Price AS decimal(18,4))),0) FROM " & mServerD & "Transaction_Details d " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = d.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "'")

    DstVATSum = ExecuteScalarCur(Cn, _
        "SELECT ISNULL(SUM(CAST(v.Valu AS decimal(18,4))),0) FROM " & mServerD & "TransactionValueAdded v " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = v.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "'")

    DstTPay = ExecuteScalarCur(Cn, _
        "SELECT ISNULL(SUM(CAST(p.Value AS decimal(18,4))),0) FROM " & mServerD & "TblTransactionPayments p " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = p.Transaction_ID " & _
        "WHERE t.SessionCode='" & SessionCode & "' AND t.Transaction_Type IN (21,9)")

    DstSPay = ExecuteScalarCur(Cn, _
        "SELECT ISNULL(SUM(CAST(s.Value AS decimal(18,4))),0) FROM " & mServerD & "TblSalesPayment s " & _
        "JOIN " & mServerD & "Transactions t ON t.Transaction_ID = s.TransID " & _
        "WHERE t.SessionCode='" & SessionCode & "' AND t.Transaction_Type IN (21,9)")

    WritePhaseLog "RECON COUNTS", _
        "SrcHeads=" & SrcHeads & _
        " | DstHeads=" & DstHeads & _
        " | SrcDet=" & SrcDet & _
        " | DstDet=" & DstDet & _
        " | SrcVAT=" & SrcVAT & _
        " | DstVAT=" & DstVAT & _
        " | SrcPay=" & SrcPay & _
        " | DstPay=" & DstPay & _
        " | SrcPay2=" & SrcPay2 & _
        " | DstPay2=" & DstPay2

    WritePhaseLog "RECON SUMS", _
        "SrcQty=" & CStr(srcQty) & _
        " | DstQty=" & CStr(dstQty) & _
        " | SrcAmount=" & CStr(SrcAmount) & _
        " | DstAmount=" & CStr(DstAmount) & _
        " | SrcVATSum=" & CStr(SrcVATSum) & _
        " | DstVATSum=" & CStr(DstVATSum) & _
        " | SrcTPay=" & CStr(SrcTPay) & _
        " | DstTPay=" & CStr(DstTPay) & _
        " | SrcSPay=" & CStr(SrcSPay) & _
        " | DstSPay=" & CStr(DstSPay)

    epsQty = 0.0001
    epsMoney = 0.01

    If (Abs(srcQty - dstQty) > epsQty) _
       Or (Abs(SrcAmount - DstAmount) > epsMoney) _
       Or (Abs(SrcVATSum - DstVATSum) > epsMoney) _
       Or (Abs(SrcTPay - DstTPay) > epsMoney) _
       Or (Abs(SrcSPay - DstSPay) > epsMoney) Then

        ReconcileMsg = "Checksum mismatch" & vbCrLf & _
                       "Qty: " & FormatNumber(srcQty, 6) & " / " & FormatNumber(dstQty, 6) & vbCrLf & _
                       "Amount: " & CStr(SrcAmount) & " / " & CStr(DstAmount) & vbCrLf & _
                       "VAT: " & CStr(SrcVATSum) & " / " & CStr(DstVATSum) & vbCrLf & _
                       "TblTransactionPayments: " & CStr(SrcTPay) & " / " & CStr(DstTPay) & vbCrLf & _
                       "TblSalesPayment: " & CStr(SrcSPay) & " / " & CStr(DstSPay)

        WriteLog "Checksum mismatch", ReconcileMsg
        Err.Raise vbObjectError + 779, "Command14_Click", ReconcileMsg
    End If

    If (DstHeads <> SrcHeads) Or (DstDet <> SrcDet) Or (DstVAT <> SrcVAT) Or (DstPay <> SrcPay) Or (DstPay2 <> SrcPay2) Then

        ReconcileMsg = "Reconcile mismatch" & vbCrLf & _
                       "Heads: " & SrcHeads & " / " & DstHeads & vbCrLf & _
                       "Details: " & SrcDet & " / " & DstDet & vbCrLf & _
                       "VAT Rows: " & SrcVAT & " / " & DstVAT & vbCrLf & _
                       "TblTransactionPayments Rows: " & SrcPay & " / " & DstPay & vbCrLf & _
                       "TblSalesPayment Rows: " & SrcPay2 & " / " & DstPay2

        WriteLog "Reconcile mismatch", ReconcileMsg
        Err.Raise vbObjectError + 778, "Command14_Click", ReconcileMsg
    End If

    '========================
    ' Final consistency check
    '========================
    CurrentStage = "Final session consistency"

    srcCount = ExecuteScalarLng(POSCn, _
        "SELECT COUNT(*) FROM Transactions WHERE IsNull(Copied,0)=0 AND SessionCode='" & SessionCode & "'")

    dstCount = ExecuteScalarLng(Cn, _
        "SELECT COUNT(*) FROM " & mServerD & "Transactions WHERE SessionCode='" & SessionCode & "'")

    If dstCount <> srcCount Then
        ReconcileMsg = "Session count mismatch before marking Copied" & vbCrLf & _
                       "Source=" & srcCount & vbCrLf & _
                       "Destination=" & dstCount
        WriteLog "Session mismatch", ReconcileMsg
        Err.Raise vbObjectError + 777, "Command14_Click", ReconcileMsg
    End If

    '========================
    ' Mark copied on source
    '========================
    CurrentStage = "Mark source as copied"

    UpdateTransferCaption " „ «· Õﬁﬁ° Ã«—Ì «⁄ „«œ «·›Ê« Ì— ﬂ„‰ﬁÊ·…", recCounter, TotalInvoices, SessionCode, mTimeStart

    LastSQL = "UPDATE Transactions SET Copied = 1, SessionCode = NULL WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "'"
    POSCn.Execute LastSQL

    Cn.CommitTrans
    inTx = False

    '========================
    ' Success UI
    '========================
    UpdateTransferCaption " „ «·‰ﬁ· »‰Ã«Õ", TotalInvoices, TotalInvoices, SessionCode, mTimeStart
    WritePhaseLog "Transfer completed successfully", "SessionCode=" & SessionCode & ", Count=" & TotalInvoices

    MousePointer = vbDefault

    r = 1

    grdInfo.TextMatrix(r, 0) = "⁄œœ «·›Ê« Ì—"
    grdInfo.TextMatrix(r, 1) = CStr(SrcHeads)
    grdInfo.TextMatrix(r, 2) = CStr(DstHeads)
    grdInfo.TextMatrix(r, 3) = IIf(SrcHeads = DstHeads, "OK", "Mismatch")
    r = r + 1

    grdInfo.TextMatrix(r, 0) = "⁄œœ ”ÿÊ— «· ›«’Ì·"
    grdInfo.TextMatrix(r, 1) = CStr(SrcDet)
    grdInfo.TextMatrix(r, 2) = CStr(DstDet)
    grdInfo.TextMatrix(r, 3) = IIf(SrcDet = DstDet, "OK", "Mismatch")
    r = r + 1

    grdInfo.TextMatrix(r, 0) = "≈Ã„«·Ì ﬁÌ„… «· ›«’Ì·"
    grdInfo.TextMatrix(r, 1) = CStr(SrcAmount)
    grdInfo.TextMatrix(r, 2) = CStr(DstAmount)
    grdInfo.TextMatrix(r, 3) = IIf(Abs(SrcAmount - DstAmount) <= epsMoney, "OK", "Mismatch")
    r = r + 1

    grdInfo.TextMatrix(r, 0) = "⁄œœ ”Ã·«  TransactionPayments"
    grdInfo.TextMatrix(r, 1) = CStr(SrcPay)
    grdInfo.TextMatrix(r, 2) = CStr(DstPay)
    grdInfo.TextMatrix(r, 3) = IIf(SrcPay = DstPay, "OK", "Mismatch")
    r = r + 1

    grdInfo.TextMatrix(r, 0) = "⁄œœ ”Ã·«  SalesPayment"
    grdInfo.TextMatrix(r, 1) = CStr(SrcPay2)
    grdInfo.TextMatrix(r, 2) = CStr(DstPay2)
    grdInfo.TextMatrix(r, 3) = IIf(SrcPay2 = DstPay2, "OK", "Mismatch")
    r = r + 1

    WritePhaseLog "Transfer completed successfully with reconcile", _
                  "SessionCode=" & SessionCode & _
                  " | SrcHeads=" & SrcHeads & _
                  " | DstHeads=" & DstHeads

    Call SaveSyncLog( _
        POSCn, Cn, SessionCode, direction, kind, _
        mTimeStart, Now, DateDiff("s", mTimeStart, Now), _
        POSlServer.Text, POSDb, SysSQLServerName, ServerDb, _
        BranchID, GetQuery, BatchSize, FetchSize, _
        SrcHeads, DstHeads, SrcDet, DstDet, SrcVAT, DstVAT, _
        SrcPay, DstPay, SrcPay2, DstPay2, _
        SrcAmount, DstAmount, SrcVATSum, DstVATSum, _
        SrcTPay, DstTPay, SrcSPay, DstSPay, _
        True, "" _
    )

    POSname_Change

EndSub:
    On Error Resume Next

    MousePointer = vbDefault
    lblWait.Visible = False

    If Not rsSer Is Nothing Then
        If rsSer.State = adStateOpen Then rsSer.Close
        Set rsSer = Nothing
    End If

    If Not rsPayments Is Nothing Then
        If rsPayments.State = adStateOpen Then rsPayments.Close
        Set rsPayments = Nothing
    End If

    If Not rsValueAdded Is Nothing Then
        If rsValueAdded.State = adStateOpen Then rsValueAdded.Close
        Set rsValueAdded = Nothing
    End If

    If Not rsDetails Is Nothing Then
        If rsDetails.State = adStateOpen Then rsDetails.Close
        Set rsDetails = Nothing
    End If

    If Not rsTrans Is Nothing Then
        If rsTrans.State = adStateOpen Then rsTrans.Close
        Set rsTrans = Nothing
    End If

    If Not rsCnt Is Nothing Then
        If rsCnt.State = adStateOpen Then rsCnt.Close
        Set rsCnt = Nothing
    End If

    If Not POSCn Is Nothing Then
        If POSCn.State = adStateOpen Then POSCn.Close
        Set POSCn = Nothing
    End If

    Exit Sub
ErrorHandler:
    Dim vErrNumber As Long
    Dim vErrDescription As String
    Dim vErrSource As String
    Dim vErrLine As Long
    Dim iErr As Integer
    Dim AdoErrText As String
    Dim PauseBeforeRollback As Boolean
    Dim ClearSourceSql As String

    '========================================
    ' «Õ›Ÿ «·Œÿ√ «·√’·Ì ›Ê—« ﬁ»· On Error Resume Next
    '========================================
    vErrNumber = Err.Number
    vErrDescription = Err.Description
    vErrSource = Err.Source
    vErrLine = Erl

    On Error Resume Next

    PauseBeforeRollback = False   ' Œ·ÌÂ True ·Ê ⁄«Ì“  Êﬁ› ﬁ»· «·‹ Rollback ··›Õ’ «·ÌœÊÌ

    UpdateTransferCaption "ÕœÀ Œÿ√ √À‰«¡ «·‰ﬁ· - Ã«—Ì «· —«Ã⁄", recCounter, TotalInvoices, SessionCode, mTimeStart

    '========================================
    ' Snapshot ··ÊÃÂ… ·Ê ﬂ«‰ ›ÌÂ DestID
    '========================================
    If Trim$(currentDestTransactionID) <> "" Then
        GetOneInvoiceSummary Cn, CDbl(currentDestTransactionID), True, _
                             OneDstDetCnt, OneDstVATCnt, OneDstPayCnt, OneDstPay2Cnt, _
                             OneDstQty, OneDstAmount, OneDstVAT, OneDstPay, OneDstPay2

        WriteLog "DEST SUMMARY ON ERROR", _
                 "DestID=" & currentDestTransactionID & _
                 " | DstDet=" & OneDstDetCnt & _
                 " | DstVAT=" & OneDstVATCnt & _
                 " | DstTPay=" & OneDstPayCnt & _
                 " | DstSPay=" & OneDstPay2Cnt & _
                 " | DstQty=" & CStr(OneDstQty) & _
                 " | DstAmount=" & CStr(OneDstAmount) & _
                 " | DstVATValue=" & CStr(OneDstVAT) & _
                 " | DstTPayValue=" & CStr(OneDstPay) & _
                 " | DstSPayValue=" & CStr(OneDstPay2)
    End If

    '========================================
    ' ”Ã· «·Œÿ√ «·√”«”Ì
    '========================================
    WritePhaseLog "Command14_Click ERROR", _
                  "Err=" & CStr(vErrNumber) & _
                  " | Desc=" & vErrDescription & _
                  " | Source=" & vErrSource & _
                  " | Line=" & CStr(vErrLine) & _
                  " | Stage=" & CurrentStage & _
                  " | InvoiceNo=" & CurrentInvoiceNo & _
                  " | SrcID=" & CStr(CurrentSrcTransactionID) & _
                  " | DestID=" & currentDestTransactionID & _
                  " | SessionCode=" & SessionCode

    WriteLog "ERROR DETAILS", _
             "Err.Number=" & CStr(vErrNumber) & _
             " | Err.Description=" & vErrDescription & _
             " | Err.Source=" & vErrSource & _
             " | Erl=" & CStr(vErrLine) & _
             " | Stage=" & CurrentStage & _
             " | InvoiceNo=" & CurrentInvoiceNo & _
             " | SrcID=" & CStr(CurrentSrcTransactionID) & _
             " | DestID=" & currentDestTransactionID & _
             " | SessionCode=" & SessionCode

    If LastSQL <> "" Then
        WriteLog "LAST SQL", LastSQL
    End If

    '========================================
    ' «· ﬁÿ ADO Errors „‰ Cn
    '========================================
    AdoErrText = ""

    If Not Cn Is Nothing Then
        If Cn.Errors.Count > 0 Then
            For iErr = 0 To Cn.Errors.Count - 1
                WriteLog "ADO Cn Error[" & CStr(iErr) & "]", _
                         "Number=" & CStr(Cn.Errors(iErr).Number) & _
                         " | NativeError=" & CStr(Cn.Errors(iErr).NativeError) & _
                         " | SQLState=" & Cn.Errors(iErr).SQLState & _
                         " | Source=" & Cn.Errors(iErr).Source & _
                         " | Desc=" & Cn.Errors(iErr).Description

                If AdoErrText = "" Then
                    AdoErrText = "Cn Error -> Number=" & CStr(Cn.Errors(iErr).Number) & _
                                 ", NativeError=" & CStr(Cn.Errors(iErr).NativeError) & _
                                 ", SQLState=" & Cn.Errors(iErr).SQLState & _
                                 ", Desc=" & Cn.Errors(iErr).Description
                End If
            Next iErr
        Else
            WriteLog "ADO Cn Error", "No provider errors"
        End If
    End If

    '========================================
    ' «· ﬁÿ ADO Errors „‰ POSCn
    '========================================
    If Not POSCn Is Nothing Then
        If POSCn.Errors.Count > 0 Then
            For iErr = 0 To POSCn.Errors.Count - 1
                WriteLog "ADO POSCn Error[" & CStr(iErr) & "]", _
                         "Number=" & CStr(POSCn.Errors(iErr).Number) & _
                         " | NativeError=" & CStr(POSCn.Errors(iErr).NativeError) & _
                         " | SQLState=" & POSCn.Errors(iErr).SQLState & _
                         " | Source=" & POSCn.Errors(iErr).Source & _
                         " | Desc=" & POSCn.Errors(iErr).Description

                If AdoErrText = "" Then
                    AdoErrText = "POSCn Error -> Number=" & CStr(POSCn.Errors(iErr).Number) & _
                                 ", NativeError=" & CStr(POSCn.Errors(iErr).NativeError) & _
                                 ", SQLState=" & POSCn.Errors(iErr).SQLState & _
                                 ", Desc=" & POSCn.Errors(iErr).Description
                End If
            Next iErr
        Else
            WriteLog "ADO POSCn Error", "No provider errors"
        End If
    End If

    '========================================
    ' ·Ê «·Œÿ√ «·√’·Ì ›«÷Ì ·ﬂ‰ ⁄‰œ‰« ADO Error
    '========================================
    If Trim$(vErrDescription) = "" And Trim$(AdoErrText) <> "" Then
        vErrDescription = AdoErrText
    End If

    '========================================
    ' «Œ Ì«—Ì: Êﬁ› ﬁ»· «·‹ Rollback
    '========================================
    If PauseBeforeRollback Then
        WriteLog "BEFORE ROLLBACK PAUSE", _
                 "SessionCode=" & SessionCode & _
                 " | Stage=" & CurrentStage & _
                 " | InvoiceNo=" & CurrentInvoiceNo & _
                 " | SrcID=" & CStr(CurrentSrcTransactionID) & _
                 " | DestID=" & currentDestTransactionID

        MsgBox " „ ≈Ìﬁ«› «· ‰›Ì– ﬁ»· Rollback." & vbCrLf & vbCrLf & _
               "SessionCode = " & SessionCode & vbCrLf & _
               "Stage = " & CurrentStage & vbCrLf & _
               "InvoiceNo = " & CurrentInvoiceNo & vbCrLf & _
               "SrcID = " & CStr(CurrentSrcTransactionID) & vbCrLf & _
               "DestID = " & currentDestTransactionID & vbCrLf & vbCrLf & _
               "«›Õ’ «·¬‰ «·”Ì—›— À„ «÷€ÿ OK ·Ì „  ‰›Ì– Rollback.", _
               vbCritical + vbOKOnly, "DEBUG BEFORE ROLLBACK"
    End If

    '========================================
    ' Rollback
    '========================================
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            If inTx Then
                Cn.RollbackTrans
                inTx = False
                WriteLog "ROLLBACK", "Done"
            End If
        End If
    End If

    '========================================
    '  ‰ŸÌ› SessionCode „‰ «·„’œ—
    '========================================
    If Not POSCn Is Nothing Then
        If POSCn.State = adStateOpen Then
            ClearSourceSql = "UPDATE Transactions " & _
                             "SET Copied = NULL, SessionCode = NULL " & _
                             "WHERE IsNull(Copied,0)=0 AND SessionCode = '" & SessionCode & "'"
            WriteLog "CLEAR SOURCE SESSION", ClearSourceSql
            POSCn.Execute ClearSourceSql
        End If
    End If

    '========================================
    ' SaveSyncLog
    '========================================
    Call SaveSyncLog( _
        POSCn, Cn, SessionCode, direction, kind, _
        mTimeStart, Now, DateDiff("s", mTimeStart, Now), _
        POSlServer.Text, POSDb, SysSQLServerName, ServerDb, _
        BranchID, GetQuery, BatchSize, FetchSize, _
        SrcHeads, DstHeads, SrcDet, DstDet, SrcVAT, DstVAT, _
        SrcPay, DstPay, SrcPay2, DstPay2, _
        SrcAmount, DstAmount, SrcVATSum, DstVATSum, _
        SrcTPay, DstTPay, SrcSPay, DstSPay, _
        False, vErrDescription)

    '========================================
    ' —”«·… Ê«÷Õ… ··„” Œœ„
    '========================================
    MsgErr = "ÕœÀ Œÿ√ √À‰«¡ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄« ." & vbCrLf & vbCrLf & _
             "«·„—Õ·…: " & CurrentStage & vbCrLf & _
             "—ﬁ„ «·›« Ê—…: " & CurrentInvoiceNo & vbCrLf & _
             "—ﬁ„ «·Õ—ﬂ… ›Ì «·„’œ—: " & CStr(CurrentSrcTransactionID) & vbCrLf & _
             "—ﬁ„ «·Õ—ﬂ… ›Ì «·”Ì—›—: " & currentDestTransactionID & vbCrLf & _
             "—ﬁ„ «·Œÿ√: " & CStr(vErrNumber) & vbCrLf & _
             "«·„’œ—: " & vErrSource & vbCrLf & _
             "«·”ÿ—: " & CStr(vErrLine) & vbCrLf & _
             "«·Ê’›: " & vErrDescription

    If ReconcileMsg <> "" Then
        MsgErr = MsgErr & vbCrLf & vbCrLf & _
                 " ›«’Ì· «·„ÿ«»ﬁ…:" & vbCrLf & ReconcileMsg
    End If

    If LastSQL <> "" Then
        MsgErr = MsgErr & vbCrLf & vbCrLf & _
                 "¬Œ— SQL  „  ”ÃÌ·Â ›Ì BatchLog.txt"
    End If

    lblWait.Visible = True
    lblWait.Caption = "›‘· «·‰ﬁ·: " & vErrDescription

    frmPopup.ShowMessage MsgErr, vbCritical

    GoTo EndSub
End Sub

Private Sub GetOneInvoiceSummary(ByVal Conn As ADODB.Connection, _
                                 ByVal TransactionID As Double, _
                                 ByVal IsDestination As Boolean, _
                                 ByRef DetCnt As Long, _
                                 ByRef VATCnt As Long, _
                                 ByRef PayCnt As Long, _
                                 ByRef Pay2Cnt As Long, _
                                 ByRef QtySum As Double, _
                                 ByRef AmountSum As Currency, _
                                 ByRef VATSum As Currency, _
                                 ByRef PaySum As Currency, _
                                 ByRef Pay2Sum As Currency)

    Dim prefix As String

    If IsDestination Then
        prefix = "dbo."
    Else
        prefix = ""
    End If

    DetCnt = ExecuteScalarLng(Conn, _
        "SELECT COUNT(*) FROM " & prefix & "Transaction_Details WHERE Transaction_ID=" & TransactionID)

    VATCnt = ExecuteScalarLng(Conn, _
        "SELECT COUNT(*) FROM " & prefix & "TransactionValueAdded WHERE Transaction_ID=" & TransactionID)

    PayCnt = ExecuteScalarLng(Conn, _
        "SELECT COUNT(*) FROM " & prefix & "TblTransactionPayments WHERE Transaction_ID=" & TransactionID)

    Pay2Cnt = ExecuteScalarLng(Conn, _
        "SELECT COUNT(*) FROM " & prefix & "TblSalesPayment WHERE TransID=" & TransactionID)

    QtySum = ExecuteScalarDbl(Conn, _
        "SELECT ISNULL(SUM(CAST(Quantity AS float)),0) FROM " & prefix & "Transaction_Details WHERE Transaction_ID=" & TransactionID)

    AmountSum = ExecuteScalarCur(Conn, _
        "SELECT ISNULL(SUM(CAST(Quantity * Price AS decimal(18,4))),0) FROM " & prefix & "Transaction_Details WHERE Transaction_ID=" & TransactionID)

    VATSum = ExecuteScalarCur(Conn, _
        "SELECT ISNULL(SUM(CAST(Valu AS decimal(18,4))),0) FROM " & prefix & "TransactionValueAdded WHERE Transaction_ID=" & TransactionID)

    PaySum = ExecuteScalarCur(Conn, _
        "SELECT ISNULL(SUM(CAST(Value AS decimal(18,4))),0) FROM " & prefix & "TblTransactionPayments WHERE Transaction_ID=" & TransactionID)

    Pay2Sum = ExecuteScalarCur(Conn, _
        "SELECT ISNULL(SUM(CAST(Value AS decimal(18,4))),0) FROM " & prefix & "TblSalesPayment WHERE TransID=" & TransactionID)

End Sub


Private Function ExecuteScalarDbl(ByVal Conn As ADODB.Connection, ByVal SQLText As String) As Double
    Dim rs As ADODB.Recordset

    On Error GoTo EH

    Set rs = Conn.Execute(SQLText)

    If rs.EOF Then
        ExecuteScalarDbl = 0#
    ElseIf IsNull(rs.Fields(0).Value) Then
        ExecuteScalarDbl = 0#
    Else
        ExecuteScalarDbl = CDbl(rs.Fields(0).Value)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    ExecuteScalarDbl = 0#
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
End Function

Private Function ExecuteScalarCur(ByVal Conn As ADODB.Connection, ByVal SQLText As String) As Currency
    Dim rs As ADODB.Recordset

    On Error GoTo EH

    Set rs = Conn.Execute(SQLText)

    If rs.EOF Then
        ExecuteScalarCur = CCur(0)
    ElseIf IsNull(rs.Fields(0).Value) Then
        ExecuteScalarCur = CCur(0)
    Else
        ExecuteScalarCur = CCur(rs.Fields(0).Value)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    ExecuteScalarCur = CCur(0)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
End Function
Private Sub UpdateTransferCaption(ByVal Msg As String, _
                                  Optional ByVal CurValue As Long = 0, _
                                  Optional ByVal TotalValue As Long = 0, _
                                  Optional ByVal SessionCode As String = "", _
                                  Optional ByVal StartTime As Date = 0)

    Dim elapsedSec As Long, elapsedMin As Long, elapsedRemSec As Long
    Dim s As String

    If StartTime <> 0 Then
        elapsedSec = DateDiff("s", StartTime, Now)
        elapsedMin = elapsedSec \ 60
        elapsedRemSec = elapsedSec Mod 60
    End If

    s = Msg

    If TotalValue > 0 Then
        s = s & " | " & CurValue & "/" & TotalValue
        s = s & " (" & Format((CurValue / TotalValue) * 100, "0") & "%)"
    End If

    If StartTime <> 0 Then
        s = s & " | «·Êﬁ : " & elapsedMin & " œ " & elapsedRemSec & " À"
    End If

    If SessionCode <> "" Then
        s = s & " | Session: " & SessionCode
    End If

    lblWait.Visible = True
    lblWait.Caption = s
    DoEvents
End Sub

Private Sub WritePhaseLog(ByVal PhaseName As String, Optional ByVal ExtraText As String = "")
    On Error Resume Next
    WriteLog PhaseName, ExtraText
End Sub

Private Function SafePercent(ByVal CurValue As Long, ByVal TotalValue As Long) As String
    If TotalValue <= 0 Then
        SafePercent = "0%"
    Else
        SafePercent = Format((CurValue / TotalValue) * 100, "0") & "%"
    End If
End Function
Sub WriteLog(Msg As String, Optional SQLText As String = "")
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\BatchLog.txt" For Append As #f
    Print #f, Now & " - " & Msg
    If SQLText <> "" Then Print #f, "    SQL: " & SQLText
    Close #f
End Sub


Function SQLDate(d As Date, Optional IncludeTime As Boolean = False) As String
    If IncludeTime Then
        SQLDate = "'" & Format(d, "yyyy-mm-dd hh:nn:ss") & "'"
    Else
        SQLDate = "'" & Format(d, "yyyy-mm-dd") & "'"
    End If
End Function


Private Sub CommandTest_Click()

    Dim POSConnection As New ADODB.Connection
    'Dim Cn As New ADODB.Connection
    Dim rsTrans As New ADODB.Recordset
    Dim BatchSize As Integer, recCounter As Integer
    Dim insertValues As String, sql As String

    ' «› Õ « ’«· «·‰ﬁÿ… POS
'    POSConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=xxx;Persist Security Info=True;User ID=xxx;Initial Catalog=POSDb;Data Source=POSServer"
'    POSConnection.Open
    
    ' «› Õ « ’«· «·”Ì—›— «·„—ﬂ“Ì
'    Cn.ConnectionString = "Provider=SQLOLEDB.1;Password=xxx;Persist Security Info=True;User ID=xxx;Initial Catalog=CentralDB;Data Source=CentralServerStaticIP"
'    Cn.Open


  Dim rsOptions As New ADODB.Recordset


    ' ≈⁄œ«œ „ €Ì—«  «·—»ÿ
    Dim mPosD As String, mServerD As String
    ' ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ…:
    mPosD = "[" & POSlServer.Text & "]." & POSDb & ".dbo."
    ' ﬁ«⁄œ… »Ì«‰«  «·”Ì—›— «·»⁄Ìœ:
    mServerD = "[RemoteServer10]." & mDBPOSName & ".dbo."

    ' ≈⁄œ«œ « ’«· «·‰ﬁÿ… (POSConnection)
    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If POSServer = "" Then POSServer = POSlServer
        If SysSQLServerType = 1 Then
            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer
        ElseIf SysSQLServerType = 2 Then
             If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & POSDb & _
                                     ";Data Source=" & POSlServer & ";Port=1433"
              Else
                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                     ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
            End If
        End If
        .Open
    End With
    

    ' «› Õ Recordset „‰ POS ··»Ì«‰«  «·„ÿ·Ê» ‰ﬁ·Â«
    rsTrans.Open "SELECT TOP 100 Transaction_ID, Transaction_Date, Transaction_Type FROM Transactions WHERE Copied IS NULL", _
                 POSConnection, adOpenForwardOnly, adLockReadOnly

    If rsTrans.EOF Then
        frmPopup.ShowMessage "·«  ÊÃœ »Ì«‰«  ··‰ﬁ·"
        Exit Sub
    End If

    ' «·„ €Ì—« 
    BatchSize = 50
    recCounter = 0
    insertValues = ""

    On Error GoTo errHandler
Cn.CursorLocation = adUseServer

    ' «»œ√ Transaction ÕﬁÌﬁÌ…
    Cn.BeginTrans

    ' Õ·ﬁ…  Ã„Ì⁄ «·»Ì«‰« 
    Do While Not rsTrans.EOF

        ' ÃÂ“ Ã„·… «·ﬁÌ„ ·ﬂ· ”Ã·
        insertValues = insertValues & "(" & _
                       rsTrans("Transaction_ID") & ", " & _
                       SQLDate(rsTrans("Transaction_Date"), True) & ", " & _
                       rsTrans("Transaction_Type") & "),"

        recCounter = recCounter + 1

        ' ‰›– batch ⁄‰œ„«  ’· ··ÕÃ„ «·„ÿ·Ê»
        If recCounter Mod BatchSize = 0 Then
            insertValues = Left(insertValues, Len(insertValues) - 1)

            sql = "INSERT INTO Transactions (Transaction_ID, Transaction_Date, Transaction_Type) VALUES " & insertValues

            ' ‰›– Ã„·… SQL
            Cn.Execute sql

            ' √⁄œ ÷»ÿ insertValues
            insertValues = ""
        End If

        rsTrans.MoveNext
    Loop

    ' ‰›– √Ì »Ì«‰«  „ »ﬁÌ…
    If insertValues <> "" Then
        insertValues = Left(insertValues, Len(insertValues) - 1)
        sql = "INSERT INTO Transactions (Transaction_ID, Transaction_Date, Transaction_Type) VALUES " & insertValues
        Cn.Execute sql
    End If

    ' Õœ¯ˆÀ «·»Ì«‰«  ›Ì POS √‰Â« „‰ﬁÊ·…
    POSConnection.Execute "UPDATE Transactions SET Copied = 1 WHERE Copied IS NULL"

    ' «⁄ „œ «·„⁄«„·… ⁄·Ï «·”Ì—›—
    Cn.CommitTrans

    frmPopup.ShowMessage " „ ‰ﬁ· «·»Ì«‰«  »‰Ã«Õ"

Exit Sub

errHandler:
    ' ›Ì Õ«·… «·Œÿ√° ﬁ„ »⁄„· Rollback
    Cn.RollbackTrans
    frmPopup.ShowMessage "ÕœÀ Œÿ√: " & Err.Description, vbCritical

End Sub

' œ«·… „”«⁄œ… · ÕÊÌ· «· «—ÌŒ ≈·Ï ’Ì€… SQL
'Function SQLDate(d As Date, Optional IncludeTime As Boolean = False) As String
'    If IncludeTime Then
'        SQLDate = "'" & Format(d, "yyyy-mm-dd hh:nn:ss") & "'"
'    Else
'        SQLDate = "'" & Format(d, "yyyy-mm-dd") & "'"
'    End If
'End Function


Private Sub Command15_Click()

'On Error GoTo ErrTrap

    ' «· √ﬂœ „‰ ÊÃÊœ ‰ﬁÿ… „ ’·…
    If POSlServer.Text = "" Then
        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
        Exit Sub
    End If
    If ConnectionFirst = False Then Exit Sub

    lblWait.Visible = True
    
    Dim TransactionIDs As String
TransactionIDs = ""
DoEvents
    ' «·Õ’Ê· ⁄·Ï «·ŒÌ«—«  „‰ TblOptions
    Dim rsOptions As New ADODB.Recordset
    rsOptions.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    JLCodeBasedOnBranch = IIf(rsOptions("JLCodeBasedOnBranch").Value = 0 Or IsNull(rsOptions("JLCodeBasedOnBranch").Value), False, True)
    StoreDigit = IIf(IsNull(rsOptions("StoreDigit").Value), 1, rsOptions("StoreDigit").Value)
    BranchDigit = IIf(IsNull(rsOptions("BranchDigit").Value), 1, rsOptions("BranchDigit").Value)
    rsOptions.Close

    ' ≈⁄œ«œ „ €Ì—«  «·—»ÿ
    Dim mPosD As String, mServerD As String
    ' ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ…:
    mPosD = "[" & POSlServer.Text & "]." & POSDb & ".dbo."
    ' ﬁ«⁄œ… »Ì«‰«  «·”Ì—›— «·»⁄Ìœ:
    mServerD = "[RemoteServer10]." & mDBPOSName & ".dbo."

    ' ≈⁄œ«œ « ’«· «·‰ﬁÿ… (POSConnection)
    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If POSServer = "" Then POSServer = POSlServer
        If SysSQLServerType = 1 Then
            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer
        ElseIf SysSQLServerType = 2 Then
             If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & POSDb & _
                                     ";Data Source=" & POSlServer & ";Port=1433"
              Else
                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                     ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
            End If
        End If
        .Open
    End With

    '  ⁄ÌÌ‰ SessionCode („À·« »‰«¡ ⁄·Ï «· «—ÌŒ Ê«·Êﬁ )
    Dim SessionCode As String
    SessionCode = Format(Now, "yyyymmddhhmmss")
Dim CarOilChangeDate As Date, RecTime As Date

    Dim mTimeStart As Date, mEndTime As Date, ActualDeliveryDate As Date, LatestDeliveryDate As Date
    mTimeStart = Now
    txtStartTime = mTimeStart
    Text3 = "Query: " & GetQuery

    ' › Õ Recordset ··„⁄«„·«  „‰ ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ…
    Dim rsTrans As New ADODB.Recordset
    sql = "SELECT * FROM Transactions WHERE Copied IS NULL AND " & GetQuery
    rsTrans.CursorType = adOpenForwardOnly
    rsTrans.LockType = adLockReadOnly
    rsTrans.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsTrans.EOF Then GoTo EndSub
    ' „ €Ì—«   Ã„Ì⁄ «·«” ⁄·«„«  (Batch)
    Dim transBatchSQL As String, detailsBatchSQL As String, valueAddedBatchSQL As String
    Dim paymentsBatchSQL As String, paymentsBatchSQL2 As String, doubleEntryBatchSQL As String, multiPaymentBatchSQL As String, notesBatchSQL As String
    transBatchSQL = ""
    detailsBatchSQL = ""
    valueAddedBatchSQL = ""
    paymentsBatchSQL = ""
    doubleEntryBatchSQL = ""
    multiPaymentBatchSQL = ""
    notesBatchSQL = ""
    
    Dim batchThreshold As Long, recCount As Long
    batchThreshold = 50
    recCount = 0

    ' »œ¡ „⁄«„·… Ê«Õœ… ⁄·Ï ﬁ«⁄œ… «·»Ì«‰«  «·ÊÃÂ… («·”Ì—›— «·»⁄Ìœ) »«” Œœ«„ Cn
    lblWait.Caption = "Ì „ «·«‰ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄« "
    DoEvents
   ' POSConnection.BeginTrans
    Dim BeginTrans As Boolean
    BeginTrans = True
On Error GoTo ErrorHandler
Dim success As Boolean
success = False
    ' «·„—Ê— ⁄·Ï ﬂ«›… «·„⁄«„·« 
    Do While Not rsTrans.EOF
         recCount = recCount + 1

         ' ﬁ—«¡… «·ÕﬁÊ· „⁄ «·ﬁÌ„ «·«› —«÷Ì… ﬂ„« ›Ì «·ﬂÊœ «·√’·Ì
         Dim PayMentType As Long, cusID As Long, BranchID As Integer, BoxID As Long, BillBasedOn As Double
         Dim VAT As Double, VATYou As Double, NoteId As Long, Trans_DiscountType As Long
         Dim Trans_Discount As Double, TaxValue As Double, order_no As Long, SaleType As Long
         Dim TaxAddValue As Double, NetValue As Double, Transaction_NetValue As Double, DepandToConv As Long
         Dim CarTypeID As Long, OilsTypesID As Long, YearFact As Long, FixesAssetsID As Long, ColorID2 As Long
         Dim KM As Double, PPointID As Long, SupplerID As Long, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double
         Dim CarEnginoil As Double, CarGearOil As Double, InvoiceTypeCodeID As Long
         Dim storeID As Variant, userID As Variant, Emp_ID As Variant
         Dim NoteSerial As String, NoteSerial1 As String, TransactionComment As String
         Dim CashCustomerName As String, CashCustomerPhone As String
         Dim PlateNo As String, Shaseh As String, CarMeter As String
         Dim CIBAN As String
         Dim InvoiceTypeCodename As String, DocumentCurrencyCode As String, TaxCurrencyCode As String
         Dim paymentnote As String, PaymentMeansCode As String
         Dim FromTransaction_Date As Date

         PayMentType = Val(rsTrans("PaymentType").Value & "")
          FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
         cusID = Val(rsTrans("CusID").Value & "")
         storeID = Val(rsTrans("StoreID").Value & "")
         userID = Val(rsTrans("UserID").Value & "")
         Emp_ID = Val(rsTrans("Emp_ID").Value & "")
         BranchID = Val(rsTrans("BranchID").Value & "")
         BoxID = Val(rsTrans("BoxID").Value & "")
         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
          PayMentType = Val(rsTrans("PaymentType").Value & "")
         cusID = Val(rsTrans("CusID").Value & "")
         storeID = Val(rsTrans("StoreID").Value & "")
         userID = Val(rsTrans("UserID").Value & "")
         Emp_ID = Val(rsTrans("Emp_ID").Value & "")
         BranchID = Val(rsTrans("BranchID").Value & "")
         BoxID = Val(rsTrans("BoxID").Value & "")
         BillBasedOn = Val(rsTrans("BillBasedOn").Value & "")
         VAT = Val(rsTrans("VAT").Value & "")
         VATYou = Val(rsTrans("VATYou").Value & "")
         NoteSerial = rsTrans("NoteSerial").Value & ""
         NoteSerial1 = rsTrans("NoteSerial1").Value & ""
         NoteId = Val(rsTrans("NoteId").Value & "")
         FromTransaction_Type = Val(rsTrans("Transaction_Type").Value & "")
         TransactionComment = rsTrans("TransactionComment").Value & ""
         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
         FromTransaction_ID = Val(rsTrans("Transaction_ID").Value & "")
         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
         TaxValue = Val(rsTrans("TaxValue").Value & "")
         order_no = Val(rsTrans("order_no").Value & "")
         SaleType = Val(rsTrans("SaleType").Value & "")
         CashCustomerName = rsTrans("CashCustomerName").Value & ""
         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
         NetValue = Val(rsTrans("NetValue").Value & "")
         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
         PlateNo = rsTrans("PlateNo").Value & ""
         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
         YearFact = Val(rsTrans("YearFact").Value & "")
         Shaseh = rsTrans("Shaseh").Value & ""
         CarMeter = rsTrans("CarMeter").Value & ""
         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
         ColorID2 = Val(rsTrans("ColorID2").Value & "")
         KM = Val(rsTrans("KM").Value & "")
         Chasee = rsTrans("Chasee").Value & ""
         PPointID = Val(rsTrans("PPointID").Value & "")
         Phone2 = rsTrans("Phone2").Value & ""
         SupplerID = Val(rsTrans("SupplerID").Value & "")
         Ser = Val(rsTrans("Ser").Value & "")
         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
         VAT = Val(rsTrans("VAT").Value & "")
         VATYou = Val(rsTrans("VATYou").Value & "")
         NoteSerial = rsTrans("NoteSerial").Value & ""
         NoteSerial1 = rsTrans("NoteSerial1").Value & ""
         NoteId = Val(rsTrans("NoteId").Value & "")
         TransactionComment = rsTrans("TransactionComment").Value & ""
         Trans_DiscountType = Val(rsTrans("Trans_DiscountType").Value & "")
         Trans_Discount = Val(rsTrans("Trans_Discount").Value & "")
         TaxValue = Val(rsTrans("TaxValue").Value & "")
         order_no = Val(rsTrans("order_no").Value & "")
         SaleType = Val(rsTrans("SaleType").Value & "")
         CashCustomerName = rsTrans("CashCustomerName").Value & ""
         TaxAddValue = Val(rsTrans("TaxAddValue").Value & "")
         CashCustomerPhone = rsTrans("CashCustomerPhone").Value & ""
         NetValue = Val(rsTrans("NetValue").Value & "")
         Transaction_NetValue = Val(rsTrans("Transaction_NetValue").Value & "")
         DepandToConv = Val(rsTrans("DepandToConv").Value & "")
         CarTypeID = Val(rsTrans("CarTypeID").Value & "")
         PlateNo = rsTrans("PlateNo").Value & ""
         OilsTypesID = Val(rsTrans("OilsTypesID").Value & "")
         YearFact = Val(rsTrans("YearFact").Value & "")
         Shaseh = rsTrans("Shaseh").Value & ""
         CarMeter = rsTrans("CarMeter").Value & ""
         FixesAssetsID = Val(rsTrans("FixesAssetsID").Value & "")
         ColorID2 = Val(rsTrans("ColorID2").Value & "")
         KM = Val(rsTrans("KM").Value & "")
         Chasee = rsTrans("Chasee").Value & ""
         PPointID = Val(rsTrans("PPointID").Value & "")
         Phone2 = rsTrans("Phone2").Value & ""
         SupplerID = Val(rsTrans("SupplerID").Value & "")
         Ser = Val(rsTrans("Ser").Value & "")
         CarCurrentValue = Val(rsTrans("CarCurrentValue").Value & "")
         CarPrevValue = Val(rsTrans("CarPrevValue").Value & "")
         CarEnginoil = Val(rsTrans("CarEnginoil").Value & "")
         CarGearOil = Val(rsTrans("CarGearOil").Value & "")
         If Trim(rsTrans("CarOilChangeDate").Value & "") = "" Then
             CarOilChangeDate = Date
         Else
             CarOilChangeDate = rsTrans("CarOilChangeDate").Value & ""
         End If
         CIBAN = rsTrans("CIBAN").Value & ""
         'RecTime = IIf(rsTrans("RecTime").Value & "" = "", Time, rsTrans("RecTime").Value & "")
         Dim tmpRecTime As Variant
tmpRecTime = rsTrans("RecTime").Value
'If IsNull(tmpRecTime) Or Trim(CStr(tmpRecTime)) = "" Or tmpRecTime = "#12/30/1899#" Then
'    RecTime = Time
'Else
'    RecTime = tmpRecTime
'End If
'RecTime = IIf(IsNull(rsTrans("RecTime").Value), Now, rsTrans("RecTime").Value)

Dim tmpRecTimeStr As String
tmpRecTimeStr = Trim(CStr(rsTrans("RecTime").Value & ""))

If tmpRecTimeStr = "" Or tmpRecTimeStr = "30-Dec-1899" Then
    RecTime = Time
ElseIf IsDate(tmpRecTimeStr) Then
    RecTime = CDate(tmpRecTimeStr)
Else
    RecTime = Time
End If

        ' RecTime = IIf(IsNull(rsTrans("RecTime").Value) Or Trim(rsTrans("RecTime").Value & "") = "", Time, rsTrans("RecTime").Value)
         ActualDeliveryDate = IIf(rsTrans("ActualDeliveryDate").Value & "" = "", Date, rsTrans("ActualDeliveryDate").Value & "")
         LatestDeliveryDate = IIf(rsTrans("LatestDeliveryDate").Value & "" = "", Date, rsTrans("ActualDeliveryDate").Value & "")
         
         InvoiceTypeCodeID = Val(rsTrans("InvoiceTypeCodeID").Value & "")
         InvoiceTypeCodename = rsTrans("InvoiceTypeCodename").Value & ""
         DocumentCurrencyCode = rsTrans("DocumentCurrencyCode").Value & ""
         TaxCurrencyCode = rsTrans("TaxCurrencyCode").Value & ""
         paymentnote = rsTrans("paymentnote").Value & ""
         PaymentMeansCode = rsTrans("PaymentMeansCode").Value & ""
         FromTransaction_Date = rsTrans("Transaction_Date").Value

         ' ≈–« POSBillType = 0°  ⁄œÌ· NoteSerial ÊNoteId
         If Val(rsTrans("POSBillType").Value & "") = 0 Then
             NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
             NoteId = Val(new_id("Notes", "NoteID", "", True) & "")
         End If

         TransactionComment = " ›« Ê—… „‰ﬁÊ·… „‰ " & POSname.Text & "   " & _
                              "   —ﬁ„ «·›« Ê—… " & NoteSerial1

         '  Ê·Ìœ —ﬁ„ ÃœÌœ ··„⁄«„·… ⁄·Ï «·ÊÃÂ…
         Dim currentDestTransactionID As String
         currentDestTransactionID = CStr((new_id("Transactions", "Transaction_ID", "", True) + recCount))

         ' »‰«¡ «” ⁄·«„ «·≈œ—«Ã «·ﬂ«„· ··„⁄«„·… »«” Œœ«„ mServerD
         Dim transSQL As String
         Dim cleanCashCustomerName As String

'  ‰ŸÌ› CashCustomerName „‰ «·ﬁÌ„ €Ì— «·’«·Õ…
        cleanCashCustomerName = Replace(CashCustomerName, ",.'", "")

         transSQL = "INSERT INTO " & mServerD & "Transactions (" & _
            "Transaction_ID, Transaction_Date, TypeInvoice, Transaction_Serial, Transaction_Type, PaymentType, " & _
                    "CusID, StoreID, UserID, Emp_ID, BranchId, BoxID, BillBasedOn, VAT, VATYou, NoteSerial, " & _
                    "NoteSerial1, NoteId, Copied, TransactionComment, SessionCode, POSBillType, OldNoteserial1, " & _
                    "OldNoteserial, OldNoteId, OldTransaction_ID, Trans_DiscountType, Trans_Discount, TaxValue, order_no, " & _
                    "SaleType, CashCustomerName, TaxAddValue, CashCustomerPhone, last_changed, NetValue, Transaction_NetValue, " & _
                    "DepandToConv, CarTypeID, PlateNo, OilsTypesID, YearFact, Shaseh, CarMeter, FixesAssetsID, " & _
                    "ColorID2, KM, Chasee, PPointID, Phone2, SupplerID, Ser, CarCurrentValue, CarPrevValue, CarEnginoil, " & _
                    "CarGearOil, CarOilChangeDate, CIBAN, RecTime, ActualDeliveryDate, LatestDeliveryDate, " & _
                    "InvoiceTypeCodeID, InvoiceTypeCodename, DocumentCurrencyCode, TaxCurrencyCode, paymentnote, PaymentMeansCode) VALUES ("
         transSQL = transSQL & currentDestTransactionID & "," & SQLDate(FromTransaction_Date, True) & "," & _
                    Val(rsTrans("TypeInvoice").Value & "") & ",'" & rsTrans("Transaction_Serial").Value & "'," & _
                    FromTransaction_Type & "," & PayMentType & "," & Val(cusID) & "," & storeID & "," & userID & "," & _
                    Emp_ID & "," & BranchID & "," & BoxID & "," & BillBasedOn & "," & VAT & "," & VATYou & "," & _
                    "'" & NoteSerial & "'," & _
                    "'" & NoteSerial1 & "',0,1," & _
                    "'" & TransactionComment & "','" & SessionCode & "'," & Val(rsTrans("POSBillType").Value & "") & "," & _
                    "'" & rsTrans("Noteserial1").Value & "','" & Trim(rsTrans("Noteserial").Value & "") & "'," & _
                    Val(rsTrans("NoteId").Value & "") & "," & rsTrans("Transaction_ID").Value & "," & _
                    Trans_DiscountType & "," & Trans_Discount & "," & TaxValue & ",'" & order_no & "'," & _
                    SaleType & ",'" & cleanCashCustomerName & "'," & TaxAddValue & ",'" & CashCustomerPhone & "'," & _
                    SQLDate(rsTrans("last_changed").Value, True) & "," & NetValue & "," & Transaction_NetValue & "," & IIf(DepandToConv, 1, 0) & "," & _
                    CarTypeID & ",'" & PlateNo & "'," & OilsTypesID & "," & YearFact & ",'" & Shaseh & "','" & CarMeter & "'," & _
                    FixesAssetsID & "," & ColorID2 & "," & KM & ",'" & Chasee & "'," & PPointID & ",'" & Phone2 & "'," & _
                    SupplerID & "," & Ser & "," & CarCurrentValue & "," & CarPrevValue & "," & CarEnginoil & "," & _
                    CarGearOil & "," & SQLDate(CarOilChangeDate, True) & ",'" & CIBAN & "'," & SQLDate(RecTime, True) & "," & _
                    SQLDate(ActualDeliveryDate, True) & "," & SQLDate(LatestDeliveryDate, True) & "," & _
                    InvoiceTypeCodeID & ",'" & InvoiceTypeCodename & "','" & DocumentCurrencyCode & "'," & _
                    "'" & TaxCurrencyCode & "','" & paymentnote & "','" & PaymentMeansCode & "')"
         transBatchSQL = transBatchSQL & transSQL & vbCrLf


If TransactionIDs = "" Then
    TransactionIDs = CStr(currentDestTransactionID)
Else
    TransactionIDs = TransactionIDs & "," & CStr(currentDestTransactionID)
End If

         ' --- „⁄«·Ã…  ›«’Ì· «·„⁄«„·… (Transaction_Details) ---
         Dim rsDetails As New ADODB.Recordset
         sql = "SELECT * FROM Transaction_Details WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
         rsDetails.CursorType = adOpenForwardOnly
         rsDetails.LockType = adLockReadOnly
         rsDetails.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
         Do While Not rsDetails.EOF
              Dim detailSQL As String
              detailSQL = "INSERT INTO " & mServerD & "Transaction_Details (" & _
                           "Transaction_ID, Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, " & _
                          "ColorID, ItemSize, ClassId, SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, AmountH, " & _
                          "AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) VALUES ("
              detailSQL = detailSQL & currentDestTransactionID & "," & Val(rsDetails("Item_ID").Value & "") & "," & Val(rsDetails("ItemCase").Value & "") & "," & _
                          Val(rsDetails("Quantity").Value & "") & "," & Val(rsDetails("Price").Value & "") & "," & Val(rsDetails("ItemDiscountType").Value & "") & "," & _
                          Val(rsDetails("ItemDiscount").Value & "") & "," & Val(rsDetails("ShowQty").Value & "") & "," & Val(rsDetails("showPrice").Value & "") & "," & _
                          Val(rsDetails("UnitId").Value & "") & "," & Val(rsDetails("ColorID").Value & "") & "," & Val(rsDetails("ItemSize").Value & "") & "," & _
                          Val(rsDetails("ClassId").Value & "") & ",'" & SessionCode & "'," & Val(rsDetails("Vatyo").Value & "") & "," & Val(rsDetails("PumpId").Value & "") & "," & _
                          Val(rsDetails("PrevQty").Value & "") & ",'" & Trim(rsDetails("PrintName").Value & "") & "'," & Val(rsDetails("Cash").Value & "") & "," & _
                          Val(rsDetails("Mada").Value & "") & "," & Val(rsDetails("Visa").Value & "") & "," & Val(rsDetails("Deferred").Value & "") & "," & _
                          Val(rsDetails("AmountH").Value & "") & "," & Val(rsDetails("AmountHComm").Value & "") & ",'" & Trim(rsDetails("DetailsPump").Value & "") & "'," & _
                          "'" & Trim(rsDetails("Account_CodeComm").Value & "") & "','" & Trim(rsDetails("Account_Code").Value & "") & "'," & _
                          IIf(rsDetails("IsOther").Value, 1, 0) & ")"
              detailsBatchSQL = detailsBatchSQL & detailSQL & vbCrLf
              rsDetails.MoveNext
         Loop
         rsDetails.Close

         ' --- „⁄«·Ã… TransactionValueAdded ---
         Dim rsValueAdded As New ADODB.Recordset
         sql = "SELECT * FROM TransactionValueAdded WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
         rsValueAdded.CursorType = adOpenForwardOnly
         rsValueAdded.LockType = adLockReadOnly
         rsValueAdded.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
         Do While Not rsValueAdded.EOF
              Dim valueSQL As String
              valueSQL = "INSERT INTO " & mServerD & "TransactionValueAdded (" & _
                         "Transaction_ID, ItemID, Vatyo, VAT, Valu, selectd, Transaction_Type, SessionCode) VALUES ("
              valueSQL = valueSQL & currentDestTransactionID & "," & rsValueAdded("ItemID").Value & "," & rsValueAdded("Vatyo").Value & "," & _
                         rsValueAdded("Vat").Value & "," & rsValueAdded("Valu").Value & "," & rsValueAdded("selectd").Value & "," & _
                         rsValueAdded("Transaction_Type").Value & ",'" & SessionCode & "')"
              valueAddedBatchSQL = valueAddedBatchSQL & valueSQL & vbCrLf
              rsValueAdded.MoveNext
         Loop
         rsValueAdded.Close

         ' --- „⁄«·Ã… TblTransactionPayments ≈–« ﬂ«‰ ‰Ê⁄ «·„⁄«„·… 21 √Ê 9 ---
         If Val(rsTrans("Transaction_Type").Value & "") = 21 Or Val(rsTrans("Transaction_Type").Value & "") = 9 Then
             Dim rsPayments As New ADODB.Recordset
             sql = "SELECT * FROM TblTransactionPayments WHERE Transaction_ID = " & Val(rsTrans("Transaction_ID").Value & "")
             rsPayments.CursorType = adOpenForwardOnly
             rsPayments.LockType = adLockReadOnly
             rsPayments.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
             Do While Not rsPayments.EOF
                  Dim paymentSQL As String, Recorddate As Date
                  Recorddate = IIf(IsNull(rsPayments("Recorddate").Value), Now, rsPayments("Recorddate").Value)
                  paymentSQL = "INSERT INTO " & mServerD & "TblTransactionPayments (" & _
                               "Transaction_ID, boxid, Recorddate, PointID, CurrentCashireID, PaymentID, Value, CardNo, Effect, SessionCode) VALUES ("
                  paymentSQL = paymentSQL & currentDestTransactionID & "," & rsPayments("boxid").Value & "," & SQLDate(Recorddate, True) & "," & _
                               rsPayments("PointID").Value & "," & rsPayments("CurrentCashireID").Value & "," & rsPayments("PaymentID").Value & "," & _
                               rsPayments("Value").Value & ",'" & rsPayments("CardNo").Value & "'," & rsPayments("Effect").Value & ",'" & SessionCode & "')"
                  paymentsBatchSQL = paymentsBatchSQL & paymentSQL & vbCrLf
                  rsPayments.MoveNext
             Loop
             rsPayments.Close
         End If

         If Val(rsTrans("Transaction_Type").Value & "") = 21 Or Val(rsTrans("Transaction_Type").Value & "") = 9 Then
             Set rsPayments = New ADODB.Recordset
             sql = "SELECT * FROM TblSalesPayment WHERE TransID = " & Val(rsTrans("Transaction_ID").Value & "")
             rsPayments.CursorType = adOpenForwardOnly
             rsPayments.LockType = adLockReadOnly
             rsPayments.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
             Do While Not rsPayments.EOF
                  Dim paymentSQL2 As String
                  
                  paymentSQL2 = "INSERT INTO " & mServerD & "TblSalesPayment (" & _
                               "TransID,  PaymentID, Value) VALUES ("
                  paymentSQL2 = paymentSQL2 & currentDestTransactionID & ", "
                  paymentSQL2 = paymentSQL2 & Val(rsPayments("PaymentID").Value & "") & ","
                  paymentSQL2 = paymentSQL2 & Val(rsPayments("Value").Value & "") & ")"
                               
                  paymentsBatchSQL2 = paymentsBatchSQL2 & paymentSQL2 & vbCrLf
                  rsPayments.MoveNext
             Loop
             rsPayments.Close
         End If


         ' --- ≈–« POSBillType = 0 And Transaction_Type <> 42° „⁄«·Ã… Notes, DOUBLE_ENTREY_VOUCHERS Ê TblMultuPayment ---
         If Val(rsTrans("POSBillType").Value & "") = 0 And Val(rsTrans("Transaction_Type").Value & "") <> 42 Then
             ' „⁄«·Ã… Notes
             Dim noteSQL As String
             noteSQL = "INSERT INTO " & mServerD & "Notes (" & _
                       "NoteID, NoteDate, NoteType, NoteSerial, NoteSerial1, branch_no, Transaction_ID, UserID, SessionCode) VALUES ("
             noteSQL = noteSQL & NoteId & "," & SQLDate(FromTransaction_Date, True) & "," & mNoteType & ",'" & NoteSerial & "'," & _
                       "'" & NoteSerial1 & "'," & BranchID & "," & currentDestTransactionID & ",1,'" & SessionCode & "')"
             notesBatchSQL = notesBatchSQL & noteSQL & vbCrLf

             ' „⁄«·Ã… DOUBLE_ENTREY_VOUCHERS
             Dim rsDoubleEntry As New ADODB.Recordset
             sql = "SELECT * FROM DOUBLE_ENTREY_VOUCHERS WHERE Notes_ID = " & NoteId
             rsDoubleEntry.CursorType = adOpenForwardOnly
             rsDoubleEntry.LockType = adLockReadOnly
             rsDoubleEntry.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
             Do While Not rsDoubleEntry.EOF
                  Dim doubleSQL As String, DEVID As String
                  DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                  doubleSQL = "INSERT INTO " & mServerD & "DOUBLE_ENTREY_VOUCHERS (" & _
                              "Double_Entry_Vouchers_ID, DEV_ID_Line_No, Account_Code, Value, Credit_Or_Debit, " & _
                              "Double_Entry_Vouchers_Description, RecordDate, Notes_ID, branch_id, UserID, Transaction_ID, SessionCode) VALUES ("
                  doubleSQL = doubleSQL & DEVID & "," & rsDoubleEntry("DEV_ID_Line_No").Value & ",'" & rsDoubleEntry("Account_Code").Value & "'," & _
                              rsDoubleEntry("Value").Value & "," & rsDoubleEntry("Credit_Or_Debit").Value & ",'" & _
                              rsDoubleEntry("Double_Entry_Vouchers_Description").Value & "'," & SQLDate(FromTransaction_Date, True) & "," & _
                              NoteId & "," & rsDoubleEntry("branch_id").Value & ",1," & currentDestTransactionID & ",'" & SessionCode & "')"
                  doubleEntryBatchSQL = doubleEntryBatchSQL & doubleSQL & vbCrLf
                  rsDoubleEntry.MoveNext
             Loop
             rsDoubleEntry.Close

             ' „⁄«·Ã… TblMultuPayment
             Dim rsMultiPay As New ADODB.Recordset
             sql = "SELECT * FROM TblMultuPayment WHERE NoteID = " & NoteId
             rsMultiPay.CursorType = adOpenForwardOnly
             rsMultiPay.LockType = adLockReadOnly
             rsMultiPay.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
             Do While Not rsMultiPay.EOF
                  Dim multiSQL As String
                  multiSQL = "INSERT INTO " & mServerD & "TblMultuPayment (" & _
                             "NoteId, PaymentID, Value, CardNo, maxvalue, SessionCode) VALUES ("
                  multiSQL = multiSQL & rsMultiPay("NoteId").Value & "," & rsMultiPay("PaymentID").Value & "," & rsMultiPay("Value").Value & ",'" & _
                             rsMultiPay("CardNo").Value & "'," & rsMultiPay("maxvalue").Value & ",'" & SessionCode & "')"
                  multiPaymentBatchSQL = multiPaymentBatchSQL & multiSQL & vbCrLf
                  rsMultiPay.MoveNext
             Loop
             rsMultiPay.Close
         End If
        lblCount.Caption = Val(lblCount.Caption) + 1
        DoEvents
         ' «·«‰ ﬁ«· ··”Ã· «· «·Ì
         rsTrans.MoveNext
            DoEvents
         '  ‰›Ì– «·œı›⁄… ≈–« Ê’· ⁄œœ «·”Ã·«  ··Õœ «·„Õœœ
         If recCount Mod batchThreshold = 0 Then
             If transBatchSQL <> "" Then POSConnection.Execute transBatchSQL: transBatchSQL = ""
             If detailsBatchSQL <> "" Then POSConnection.Execute detailsBatchSQL: detailsBatchSQL = ""
             If valueAddedBatchSQL <> "" Then POSConnection.Execute valueAddedBatchSQL: valueAddedBatchSQL = ""
             If paymentsBatchSQL <> "" Then POSConnection.Execute paymentsBatchSQL: paymentsBatchSQL = ""
             If paymentsBatchSQL2 <> "" Then POSConnection.Execute paymentsBatchSQL2: paymentsBatchSQL2 = ""
          '   If doubleEntryBatchSQL <> "" Then POSConnection.Execute doubleEntryBatchSQL: doubleEntryBatchSQL = ""
             If multiPaymentBatchSQL <> "" Then POSConnection.Execute multiPaymentBatchSQL: multiPaymentBatchSQL = ""
            ' If notesBatchSQL <> "" Then POSConnection.Execute notesBatchSQL: notesBatchSQL = ""
         End If

    Loop
    rsTrans.Close

    '  ‰›Ì– √Ì œ›⁄«  „ »ﬁÌ…
    If transBatchSQL <> "" Then POSConnection.Execute transBatchSQL
    If detailsBatchSQL <> "" Then POSConnection.Execute detailsBatchSQL
    If valueAddedBatchSQL <> "" Then POSConnection.Execute valueAddedBatchSQL
    If paymentsBatchSQL <> "" Then POSConnection.Execute paymentsBatchSQL
    If paymentsBatchSQL2 <> "" Then POSConnection.Execute paymentsBatchSQL2
  '  If doubleEntryBatchSQL <> "" Then POSConnection.Execute doubleEntryBatchSQL
    If multiPaymentBatchSQL <> "" Then POSConnection.Execute multiPaymentBatchSQL
   ' If notesBatchSQL <> "" Then Cn.Execute notesBatchSQL

    '  ÕœÌÀ ”Ã·«  Transactions ›Ì ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ… · ⁄ÌÌ‰ Copied = 1 „⁄ SessionCode
    sql = "UPDATE " & mPosD & "Transactions SET Copied = 1, SessionCode = '" & SessionCode & "' WHERE Copied IS NULL AND " & GetQuery
    POSConnection.Execute sql

    ' ≈–« ﬂ«‰ chkRec „›⁄·°  ÕœÌÀ ÃœÊ· Notes ›Ì ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ…
    If chkRec.Value = vbChecked Then
         sql = "UPDATE " & mPosD & "Notes SET Copied = 1, SessionCode = '" & SessionCode & "' WHERE NoteType = 4 AND Copied IS NULL AND NoteDate = " & SQLDate(dbRecordDate.Value, False)
         POSConnection.Execute sql
    End If

    '  ”ÃÌ· ”Ã· ›Ì TblOffline · ÊÀÌﬁ ⁄„·Ì… «·‰ﬁ·
    Dim rsOffline As New ADODB.Recordset
    sql = "SELECT * FROM TblOffline WHERE 1 = -1"
    rsOffline.Open sql, Cn, adOpenKeyset, adLockOptimistic
    rsOffline.AddNew
    rsOffline!Recorddate = Date
    rsOffline!StartTime = mTimeStart
    mEndTime = Now
    rsOffline!EndTime = mEndTime
    rsOffline!SessionCode = SessionCode
    rsOffline!POSname = POSlServer.Text
    rsOffline!CountSalesOfeers = CountSalesOfeers
    rsOffline!CountSales = CountSales
    rsOffline!CountSalesReturn = CountSalesReturn
    rsOffline!CountPurchase = CountPurchase
    rsOffline!CountPurchaseReturn = CountPurchaseReturn
    rsOffline!CountRec = CountRec
    rsOffline.Update
    rsOffline.Close


 sql = "SELECT * FROM TblOffline WHERE 1 = -1"
    rsOffline.Open sql, POSConnection, adOpenKeyset, adLockOptimistic
    rsOffline.AddNew
    rsOffline!Recorddate = Date
    rsOffline!StartTime = mTimeStart
    mEndTime = Now
    rsOffline!EndTime = mEndTime
    rsOffline!SessionCode = SessionCode
    rsOffline!POSname = POSlServer.Text
    rsOffline!CountSalesOfeers = CountSalesOfeers
    rsOffline!CountSales = CountSales
    rsOffline!CountSalesReturn = CountSalesReturn
    rsOffline!CountPurchase = CountPurchase
    rsOffline!CountPurchaseReturn = CountPurchaseReturn
    rsOffline!CountRec = CountRec
    rsOffline.Update
    rsOffline.Close
    ' ≈‰Â«¡ «·„⁄«„·…
   ' POSConnection.CommitTrans
    BeginTrans = False
    success = True
    
    lblWait.Visible = True
    DoEvents
    txtEndTime = mEndTime
    txtCountSalesReturn = CountSalesReturn
    txtCountSales = CountSales
    txtCountSalesOfeers = CountSalesOfeers
    
    lblWait.Caption = " „ ‰ﬁ· ›Ê« Ì— «·„»Ì⁄«  »‰Ã«Õ"
EndSub:
    frmPopup.ShowMessage " „ «·‰ﬁ· »‰Ã«Õ"
    DoEvents
    Exit Sub
ErrorHandler:
If success Then
    ' ??? ??????? ????? ??? ??? ??????
    
    sql = "UPDATE " & mPosD & "Transactions SET Copied = null, SessionCode = null WHERE SessionCode = '" & SessionCode & "' and IsNull(SessionCode,'') <> ''"
    POSConnection.Execute sql
    
    
    
    If TransactionIDs <> "" Then
        Dim deleteSQL As String
        deleteSQL = "DELETE FROM Transactions WHERE Transaction_ID IN (" & TransactionIDs & ") and SessionCode = '" & SessionCode & "'"
        Cn.Execute deleteSQL
    End If
    
    
    
End If


    sql = "UPDATE " & mPosD & "Transactions SET Copied = null, SessionCode = null WHERE SessionCode = '" & SessionCode & "' and IsNull(SessionCode,'') <> ''"
    POSConnection.Execute sql
    
    
    
    If TransactionIDs <> "" Then
        
        deleteSQL = "DELETE FROM Transactions WHERE Transaction_ID IN (" & TransactionIDs & ") and SessionCode = '" & SessionCode & "'"
        Cn.Execute deleteSQL
    End If
    MsgBox "Œÿ√ «À‰«¡ «·‰ﬁ· —Ã«¡ «· Ê«’· „⁄ „”∆Ê·Ï «·‰Ÿ«„: " & Err.Description, vbCritical, "???"
Exit Sub
ErrTrap:
    If BeginTrans Then
         POSConnection.RollbackTrans
         BeginTrans = False
    End If
    If Err.Number = -2147217900 Then
         MsgBox "?C ???? ??U ??? C?E?C?CE" & vbCrLf & "??I E? CIIC? ??? U?? ?C??E", vbExclamation, App.Title
         Exit Sub
    End If
    MsgBox "Œÿ√: " & Err.Description, vbExclamation, App.Title


End Sub

'Private Sub Command2_Click()
'
'
'
''   ************************************'check items here first wael*******************
' Dim StrSQL As String
'If POSlServer.Text = "" Then
'MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
'Exit Sub
'End If
'
''
'
'  Dim mPosD As String
'            Dim mServerD As String
'            mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
'            mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
'            mServerD = "RemoteServer10"
'
'            'mServerD = "[" & SysSQLServerName & "]." & ServerDb & ".dbo."
'
'           mServerD = "[RemoteServer10]." & mDBPOSName & ".dbo."
'
''Command4_Click
'lblWait.Visible = True
'   Dim NoOFItem_POS As Double
'   Dim NoOFItem_Server As Double
'
'   Dim Rs3 As New ADODB.Recordset
'   Dim MaxItem_POS As Double
'   Dim MaxItem_Server As Double
'   'step one check item
'
'    ss = "     USE " & ServerDb & vbNewLine
'
'    Cn.Execute ss
'    ss = "USE " & POSDb & vbNewLine
'    POSConnection.Execute ss
'
'    sql = " select count (ItemID ) As NoOfitems ,max(ItemID) as MaxItemid from TblItems  "
'
'    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
'        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
'
'    End If
'    Rs3.Close
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'   ' MsgBox "Step 1"
'    If Rs3.RecordCount > 0 Then
'        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
'        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
'    End If
'    Rs3.Close
'
'   ' MsgBox "Item Server" & NoOFItem_Server
'   ' MsgBox "Item Pos" & NoOFItem_POS
'    'step 2
'   ' Exit Sub
'   lblWait.Caption = "Ì „ «·«‰  ÕœÌÀ «·«”⁄«— Ê«·„·›«  «·«”«”Ì…"
'    If 1 = 1 Then
'             'checkGroup
'        Dim NoOfGroups_pos As Double
'        Dim NoOfGroups_server As Double
'
'        Dim MaxGroupid_pos As Double
'        Dim MaxGroupidserver As Double
'
'
'        sql = " select count (GroupID ) As NoOfGroups ,max(GroupID) as MaxGroupid from Groups  "
'        Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs3.RecordCount > 0 Then
'            NoOfGroups_pos = IIf(IsNull(Rs3("NoOfGroups").Value), 0, Rs3("NoOfGroups").Value)
'            MaxGroupid_pos = IIf(IsNull(Rs3("MaxGroupid").Value), 0, Rs3("MaxGroupid").Value)
'        End If
'        Rs3.Close
'        'MsgBox "Step 2"
'
'
'        sql = " select count (GroupID ) As NoOfGroups ,max(GroupID) as MaxGroupid from Groups  "
'
'        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs3.RecordCount > 0 Then
'            NoOfGroups_server = IIf(IsNull(Rs3("NoOfGroups").Value), 0, Rs3("NoOfGroups").Value)
'            MaxGroupidserver = IIf(IsNull(Rs3("MaxGroupid").Value), 0, Rs3("MaxGroupid").Value)
'        End If
'        Rs3.Close
'
'         'MsgBox "Step 3"
'        Dim s As String
'
'        If 1 = 1 Then
'
'
'            BolFrmLoaded = True
'
'
'         Do While Rs3.State = adStateExecuting
'                DoEvents
'            Loop
'
'
'            s = ""
'
'
'
'       sql = "INSERT INTO " & mPosD & "Groups (GroupID, GroupName) " & _
'      "SELECT T2.GroupID, T2.GroupName " & _
'      "FROM [RemoteServer10].byte.dbo.Groups T2 " & _
'      "WHERE NOT EXISTS ( " & _
'      "    SELECT 1 " & _
'      "    FROM " & mPosD & "Groups Tpos " & _
'      "    WHERE Tpos.GroupID = T2.GroupID);"
'
'
'           ' Text4 = s
'           ' Exit Sub
'            POSConnection.Execute sql
'
'
'
'' ' ≈œŒ«· «·»Ì«‰«  «·ÃœÌœ… - TblUnites
''s = "INSERT INTO " & mPosD & "TblUnites " & _
''    "SELECT * " & _
''    "FROM " & mServerD & "TblUnites T2 " & _
''    "LEFT JOIN " & mPosD & "TblUnites T1 ON T2.UnitID = T1.UnitID " & _
''    "WHERE T1.UnitID IS NULL;"
'
'
'' ‰ﬁ· «·»Ì«‰«  „‰ TblUnites
''s = "INSERT INTO " & mPosD & "TblUnites (UnitID, UnitName, UnitNamee) " & _
''    "SELECT T2.UnitID, T2.UnitName, T2.UnitNamee " & _
''    "FROM " & mServerD & "TblUnites T2 " & _
''    "LEFT JOIN " & mPosD & "TblUnites T1 ON T2.UnitID = T1.UnitID " & _
''    "WHERE T1.UnitID IS NULL;"
''POSConnection.Execute s
''
''' ‰ﬁ· «·»Ì«‰«  „‰ TblItems
''s = "INSERT INTO " & mPosD & "TblItems (ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, " & _
''    "CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, " & _
''    "ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, barCodeNO, SizeID11) " & _
''    "SELECT T2.ItemID, T2.ItemCode, T2.ItemName, T2.GroupID, T2.HaveSerial, T2.LastUpdate, T2.PurchasePrice, T2.SallingPrice, T2.RequestLimit, " & _
''    "T2.CustomerPrice, T2.HaveGuarantee, T2.GuaranteeValue, T2.GuaranteeType, T2.IsArchive, T2.ItemType, T2.AssbliedItem, T2.RelatedItem, T2.ItemComment, T2.ItemCase, T2.ItemMaking, " & _
''    "T2.ItemMakingNew, T2.code, T2.Branch_NO, T2.Fullcode, T2.prifix, T2.PartNo, T2.CostPrice, T2.ItemNamee, T2.DefaultSupplier, T2.itemSerials, T2.barCodeNO, T2.SizeID11 " & _
''    "FROM " & mServerD & "TblItems T2 " & _
''    "LEFT JOIN " & mPosD & "TblItems T1 ON T2.ItemID = T1.ItemID " & _
''    "WHERE T1.ItemID IS NULL;"
''POSConnection.Execute s
''
''' ‰ﬁ· «·»Ì«‰«  „‰ TblItemsUnits
''s = "INSERT INTO " & mPosD & "TblItemsUnits (JunckID, ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, " & _
''    "MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2) " & _
''    "SELECT T2.JunckID, T2.ItemID, T2.UnitID, T2.UnitFactor, T2.SecOrder, T2.DefaultUnit, T2.UnitSalesPrice, T2.UnitPurPrice, T2.FactorByDefaultUnit, " & _
''    "T2.MinSelingPrice, T2.ForUnit, T2.MethodCalc, T2.SessionCode, T2.barCodeNo2 " & _
''    "FROM " & mServerD & "TblItemsUnits T2 " & _
''    "LEFT JOIN " & mPosD & "TblItemsUnits T1 ON T2.ItemID = T1.ItemID " & _
''    "WHERE T1.ItemID IS NULL;"
''POSConnection.Execute s
'
'
'
'
'
'
'
'
'' «· √ﬂœ „‰ √‰ «·« ’«· „› ÊÕ
'If POSConnection.State = 0 Then POSConnection.Open
'
'On Error Resume Next
'
'' ‰ﬁ· «·»Ì«‰«  „‰ TblUnites
's = "INSERT INTO " & mPosD & "TblUnites (UnitID, UnitName, UnitNamee) " & _
'    "SELECT T2.UnitID, T2.UnitName, T2.UnitNamee " & _
'    "FROM " & mServerD & "TblUnites T2 " & _
'    "LEFT JOIN " & mPosD & "TblUnites T1 ON T2.UnitID = T1.UnitID " & _
'    "WHERE T1.UnitID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblUnites: " & Err.Description
'    Err.Clear
'End If
'
'' ‰ﬁ· «·»Ì«‰«  „‰ TblItems
's = "INSERT INTO " & mPosD & "TblItems (ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, " & _
'    "CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, " & _
'    "ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, barCodeNO, SizeID11) " & _
'    "SELECT T2.ItemID, T2.ItemCode, T2.ItemName, T2.GroupID, T2.HaveSerial, T2.LastUpdate, T2.PurchasePrice, T2.SallingPrice, T2.RequestLimit, " & _
'    "T2.CustomerPrice, T2.HaveGuarantee, T2.GuaranteeValue, T2.GuaranteeType, T2.IsArchive, T2.ItemType, T2.AssbliedItem, T2.RelatedItem, T2.ItemComment, T2.ItemCase, T2.ItemMaking, " & _
'    "T2.ItemMakingNew, T2.code, T2.Branch_NO, T2.Fullcode, T2.prifix, T2.PartNo, T2.CostPrice, T2.ItemNamee, T2.DefaultSupplier, T2.itemSerials, T2.barCodeNO, T2.SizeID11 " & _
'    "FROM " & mServerD & "TblItems T2 " & _
'    "LEFT JOIN " & mPosD & "TblItems T1 ON T2.ItemID = T1.ItemID " & _
'    "WHERE T1.ItemID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblItems: " & Err.Description
'    Err.Clear
'End If
'
'' ‰ﬁ· «·»Ì«‰«  „‰ TblItemsUnits
'
'
'' ≈€·«ﬁ «·« ’«· ⁄‰œ «·«‰ Â«¡
''POSConnection.Close
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'           ' MsgBox "Step 4"
''            s = " INSERT INTO " & mPosD & "TblUnites"
''            s = s & " SELECT *"
''            s = s & " FROM   " & mServerD & "TblUnites T2"
''            s = s & " WHERE  T2.UnitID NOT IN (SELECT IsNull(Tpos.UnitID,0)"
''            s = s & "                                      FROM   " & mPosD & "TblUnites as  Tpos);"
''
''            Cn.Execute s
''
''           ' MsgBox "Step 5"
'''            s = " INSERT INTO " & mPosD & "TblItemLoc"
'''            s = s & " SELECT *"
'''            s = s & " FROM   " & mServerD & "TblItemLoc T2"
'''            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'''            s = s & "                                      FROM   " & mPosD & "TblItemLoc);"
'''
'''            Cn.Execute s
'''
''
''
''            s = " INSERT INTO " & mPosD & "TblItems"
''            s = s & " SELECT * "
''            s = s & " FROM   " & mServerD & "TblItems T2"
''            s = s & " WHERE  T2.ItemID NOT IN (SELECT IsNull(Tpos.ItemID,0)"
''            s = s & "                                      FROM   " & mPosD & "TblItems as Tpos);"
''
''
''            Cn.Execute s
''
''
''           ' MsgBox "Step 6"
''           '1
'''            MsgBox "pos" & mPosD
'''            MsgBox "Server " & mServerD
'''
'''                       s = " INSERT INTO " & mPosD & "TblItemProductLine"
'''            s = s & " SELECT *"
'''            s = s & " FROM   " & mServerD & "TblItemProductLine T2"
'''            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'''            s = s & "                                      FROM   " & mPosD & "TblItemProductLine);"
'''
'''            Cn.Execute s
'''
''
'''            s = " INSERT INTO " & mPosD & "TblItemsAttach"
'''            s = s & " SELECT *"
'''            s = s & " FROM   " & mServerD & "TblItemsAttach T2"
'''            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'''            s = s & "                                      FROM   " & mPosD & "TblItemsAttach);"
'''
'''            Cn.Execute s
'''
'''            s = " INSERT INTO " & mPosD & "ItemsPrice"
'''            s = s & " SELECT *"
'''            s = s & " FROM   " & mServerD & "ItemsPrice T2"
'''            s = s & " WHERE  T2.Item_ID NOT IN (SELECT Item_ID"
'''            s = s & "                                      FROM   " & mPosD & "ItemsPrice);"
'''
'''            Cn.Execute s
'''
'''
''
'''            s = " INSERT INTO  " & mPosD & "ItemsParts"
'''            s = s & " SELECT *"
'''            s = s & " FROM   " & mServerD & "ItemsParts T2"
'''            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'''            s = s & "                                      FROM   " & mPosD & "ItemsParts);"
'''
'''            Cn.Execute s
''
''        '    MsgBox "Step 6"
''            s = " INSERT INTO " & mPosD & "TblItemsUnits"
''            s = s & " SELECT *"
''            s = s & " FROM   " & mServerD & "TblItemsUnits T2"
''            s = s & " WHERE  T2.ItemID NOT IN (SELECT IsNull(TPos.ItemID,0)"
''            s = s & "                                      FROM   " & mPosD & "TblItemsUnits as TPos);"
''
''            Cn.Execute s
'
'            Text5 = s
'
'     '       MsgBox "Step 7"
'
'
'            'Copy  remains Groups
'            'Copy  remains Items
'            'Copy itemsunits
'
'
'            ' MsgBox " „ ‰ﬁ· »Ì«‰«  «·«’‰«›"
'             Command2.Enabled = False
'
'        End If
'     Else
'
'    '  MsgBox "    „·›   «·«’‰«› „ÕœÀ"
'    '  lblWait.Visible = False
'
'End If
'    '  ÕœÌÀ «·»Ì«‰«  - TblItemsUnits
'
'
'
'' ?????? ?? ?? ??????? ?????
'If POSConnection.State = 0 Then POSConnection.Open
'
'On Error Resume Next
'
'' ??? ???????? ?? TblPaymentType
's = "INSERT INTO " & mPosD & "TblPaymentType (PaymentID, PaymentName, PaymentNamee, Accountcom, commision, branch_no, TaxTobacco, AccTaxTobacco, IsNewCode, IsHiddenVat, IsDefault) " & _
'    "SELECT T2.PaymentID, T2.PaymentName, T2.PaymentNamee, T2.Accountcom, T2.commision, T2.branch_no, T2.TaxTobacco, T2.AccTaxTobacco, T2.IsNewCode, T2.IsHiddenVat, T2.IsDefault " & _
'    "FROM " & mServerD & "TblPaymentType T2 " & _
'    "LEFT JOIN " & mPosD & "TblPaymentType T1 ON T2.PaymentID = T1.PaymentID " & _
'    "WHERE T1.PaymentID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblPaymentType: " & Err.Description
'    Err.Clear
'End If
'DoEvents
'' ??? ???????? ?? TblPaymentUser
's = "INSERT INTO " & mPosD & "TblPaymentUser ( PaynetID, UserID) " & _
'    "SELECT  T2.PaynetID, T2.UserID " & _
'    "FROM " & mServerD & "TblPaymentUser T2 " & _
'    "LEFT JOIN TblPaymentUser T1 ON T2.id = T1.id " & _
'    "WHERE T1.id IS NULL;"
'
'
'   'POSConnection.Execute "SET IDENTITY_INSERT TblPaymentUser off"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'  '  frmPopup.ShowMessage "Error in TblPaymentUser: " & Err.Description
'    Err.Clear
'End If
''POSConnection.Execute "SET IDENTITY_INSERT TblPaymentUser Off"
'' ????? ??????? ??? ????????
'
'
'
'' ?????? ?? ?? ??????? ?????
'If POSConnection.State = 0 Then POSConnection.Open
'
'On Error Resume Next
'
'' ??? ???????? ?? BanksData
'lblWait.Caption = "Ì „ «·«‰  ÕœÌÀ „·› «·»‰Êﬂ"
's = "INSERT INTO " & mPosD & "BanksData (BankID, BankName, BankNamee, Account_Code, Account_Code1, Account_Code2, BranchId, ParetnAccount, parent_account) " & _
'    "SELECT T2.BankID, T2.BankName, T2.BankNamee, T2.Account_Code, T2.Account_Code1, T2.Account_Code2, T2.BranchId, T2.ParetnAccount, T2.parent_account " & _
'    "FROM " & mServerD & "BanksData T2 " & _
'    "LEFT JOIN " & mPosD & "BanksData T1 ON T2.BankID = T1.BankID " & _
'    "WHERE T1.BankID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in BanksData: " & Err.Description
'    Err.Clear
'End If
'lblWait.Caption = "Ì „ «·«‰  ÕœÌÀ ««·„” Œœ„Ì‰"
'DoEvents
'' ??? ???????? ?? TblUsers
's = "INSERT INTO " & mPosD & "TblUsers (UserID, UserName, PassWord, BranchId, BoxID, BankID, Empid, FixedCustomer) " & _
'    "SELECT T2.UserID, T2.UserName, T2.PassWord, T2.BranchId, T2.BoxID, T2.BankID, T2.Empid, T2.FixedCustomer " & _
'    "FROM " & mServerD & "TblUsers T2 " & _
'    "LEFT JOIN " & mPosD & "TblUsers T1 ON T2.UserID = T1.UserID " & _
'    "WHERE T1.UserID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblUsers: " & Err.Description
'    Err.Clear
'End If
'
'
'
'' ??? ???????? ?? TblEmpJobsTypes
's = "INSERT INTO " & mPosD & "TblEmpJobsTypes (JobTypeID, JobTypeName, JobTypeNamee) " & _
'    "SELECT T2.JobTypeID, T2.JobTypeName, T2.JobTypeNamee " & _
'    "FROM " & mServerD & "TblEmpJobsTypes T2 " & _
'    "LEFT JOIN " & mPosD & "TblEmpJobsTypes T1 ON T2.JobTypeID = T1.JobTypeID " & _
'    "WHERE T1.JobTypeID IS NULL;"
'POSConnection.Execute s
'
'' ?????? ?? ???????
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblEmpJobsTypes: " & Err.Description
'    Err.Clear
'End If
'lblWait.Caption = "Ì „ «·«‰  ÕœÌÀ «·„ÊŸ›Ì‰"
'DoEvents
'' ??? ???????? ?? TblEmployee
's = "INSERT INTO " & mPosD & "TblEmployee (Emp_ID, Emp_Code, Emp_Name, Nationality, dean, JobTypeID, placeWORK, DepartmentID, Emp_Salary, Emp_Salary_others, NumEkama, DateEndekamah, KafelID, NumPasp, jopstatusid, Emp_mobile, BranchId, Emp_Namee) " & _
'    "SELECT T2.Emp_ID, T2.Emp_Code, T2.Emp_Name, T2.Nationality, T2.dean, T2.JobTypeID, T2.placeWORK, T2.DepartmentID, T2.Emp_Salary, T2.Emp_Salary_others, T2.NumEkama, T2.DateEndekamah, T2.KafelID, T2.NumPasp, T2.jopstatusid, T2.Emp_mobile, T2.BranchId, T2.Emp_Namee " & _
'    "FROM " & mServerD & "TblEmployee T2 " & _
'    "LEFT JOIN " & mPosD & "TblEmployee T1 ON T2.Emp_ID = T1.Emp_ID " & _
'    "WHERE T1.Emp_ID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblEmployee: " & Err.Description
'    Err.Clear
'End If
'lblWait.Caption = "Ì „ «·«‰  ÕœÌÀ «·⁄„·«¡"
'DoEvents
'' ??? ???????? ?? TblCustemers
's = "INSERT INTO " & mPosD & "TblCustemers (CusID, CusName, CusNamee, ResponsibleContact, Cus_mobile, Type, OpenBalance, Account_Code, CityID, EmpId, Address, parent_account, prifix, Fullcode, BranchId, VATNO, CustGID) " & _
'    "SELECT T2.CusID, T2.CusName, T2.CusNamee, T2.ResponsibleContact, T2.Cus_mobile, T2.Type, T2.OpenBalance, T2.Account_Code, T2.CityID, T2.EmpId, T2.Address, T2.parent_account, T2.prifix, T2.Fullcode, T2.BranchId, T2.VATNO, T2.CustGID " & _
'    "FROM " & mServerD & "TblCustemers T2 " & _
'    "LEFT JOIN " & mPosD & "TblCustemers T1 ON T2.CusID = T1.CusID " & _
'    "WHERE T1.CusID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblCustemers: " & Err.Description
'    Err.Clear
'End If
'
'' ??? ???????? ?? TblStore
's = "INSERT INTO " & mPosD & "TblStore (StoreID, StoreName, Account_Code, Account_Code1, Account_Code2, Emp_ID, Account_Code3, linked, BranchId, Code, StoreNamee, ParetnAccount, SalesPersonId, PurchasePersonid, Account_Code0, Account_Code11, Account_Code22, Account_Code33, BoxID) " & _
'    "SELECT T2.StoreID, T2.StoreName, T2.Account_Code, T2.Account_Code1, T2.Account_Code2, T2.Emp_ID, T2.Account_Code3, T2.linked, T2.BranchId, T2.Code, T2.StoreNamee, T2.ParetnAccount, T2.SalesPersonId, T2.PurchasePersonid, T2.Account_Code0, T2.Account_Code11, T2.Account_Code22, T2.Account_Code33, T2.BoxID " & _
'    "FROM " & mServerD & "TblStore T2 " & _
'    "LEFT JOIN " & mPosD & "TblStore T1 ON T2.StoreID = T1.StoreID " & _
'    "WHERE T1.StoreID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblStore: " & Err.Description
'    Err.Clear
'End If
'
'' ????? ??????? ??? ????????
'
's = "INSERT INTO " & mPosD & "TblItemsUnits (JunckID, ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, " & _
'    "MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2) " & _
'    "SELECT T2.JunckID, T2.ItemID, T2.UnitID, T2.UnitFactor, T2.SecOrder, T2.DefaultUnit, T2.UnitSalesPrice, T2.UnitPurPrice, T2.FactorByDefaultUnit, " & _
'    "T2.MinSelingPrice, T2.ForUnit, T2.MethodCalc, T2.SessionCode, T2.barCodeNo2 " & _
'    "FROM " & mServerD & "TblItemsUnits T2 " & _
'    "LEFT JOIN " & mPosD & "TblItemsUnits T1 ON T2.ItemID = T1.ItemID " & _
'    "WHERE T1.ItemID IS NULL;"
'POSConnection.Execute s
'If Err.Number <> 0 Then
'    frmPopup.ShowMessage "Error in TblItemsUnits: " & Err.Description
'    Err.Clear
'End If
'
'
's = "UPDATE T1 " & _
'    "SET T1.UnitSalesPrice = T2.UnitSalesPrice,T1.barCodeNo2=T2.barCodeNo2, " & _
'    "    T1.MaxSelingPrice = T2.MaxSelingPrice, " & _
'    "    T1.UnitWholeSalePrice = T2.UnitWholeSalePrice, " & _
'    "    T1.MinSelingPrice = T2.MinSelingPrice, " & _
'    "    T1.UnitPurPrice = T2.UnitPurPrice " & _
'    "FROM " & mPosD & "TblItemsUnits T1 " & _
'    "INNER JOIN " & mServerD & "TblItemsUnits T2 " & _
'    "ON T1.ItemID = T2.ItemID AND T1.UnitId = T2.UnitId;"
'POSConnection.Execute s
''MsgBox ""
'Text6 = s
''  ÕœÌÀ «·»Ì«‰«  - TblItems
's = "UPDATE T1 " & _
'    "SET T1.ItemName = T2.ItemName, " & _
'    "    T1.barCodeNO = T2.barCodeNO, " & _
'    "    T1.Code = T2.Code, " & _
'    "    T1.Fullcode = T2.Fullcode, " & _
'    "    T1.IsArchive = ISNULL(T2.IsArchive, 0) " & _
'    "FROM " & mPosD & "TblItems T1 " & _
'    "INNER JOIN " & mServerD & "TblItems T2 " & _
'    "ON T1.ItemID = T2.ItemID;"
'POSConnection.Execute s
'
'  POSConnection.Close
'   '************************************'check items here first*******************
'   lblWait.Visible = True
'lblWait.Caption = " „   ÕœÌÀ «·«”⁄«— Ê«·„·›«  «·«”«”Ì… »‰Ã«Õ"
'DoEvents
'End Sub
'
Private Sub Command4_Click()
    If ConnectionFirst = False Then
        Exit Sub
    End If
    Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If



   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   'step one check item
       
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    Text2 = ss & " " & POSConnection.ConnectionString
    sql = " select count (ItemID ) As NoOfitems ,max(ItemID) as MaxItemid from TblItems  "
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
  '  MsgBox "⁄œœ «’‰«› «·‰ﬁÿ…" & NoOFItem_POS
    Rs3.Close
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount > 0 Then
'        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
'        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
'    End If
'    Rs3.Close
'
'
    'step 2
    
End Sub

Private Sub Command5_Click()
    Dim s As String
    Dim mPosD As String
    Dim mServerD As String
     mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
     mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
     mServerD = ServerDb & ".dbo."
            
   
    
    s = " Update   " & mServerD & "Transactions Set OldNoteSerial1 = (Select NoteSerial1 From  " & mPosD & "Transactions T "
    s = s & " Where T.Transaction_ID = Transactions.OldTransaction_ID "
    s = s & "  AND T.BranchId = Transactions.BranchID"
    s = s & " and T.Transaction_Type = Transactions.Transaction_Type )"
    s = s & " Where Transactions.Transaction_Type = 21 and Transactions.BranchID = " & Val(DCboBranch.BoundText) & " and IsNull(SessionCode,'') <> ''"
    s = s & " AND Transactions.OldTransaction_ID IN (SELECT Transaction_ID FROM   " & mPosD & "Transactions T Where T.Transaction_Type = 21 )"
    Cn.Execute s
    
    frmPopup.ShowMessage " „ «·‰ﬁ·"
End Sub

Private Sub Command6_Click()
Dim s As String
Dim rsDummy As New ADODB.Recordset
Dim TransType  As Long
Dim mYear As Integer
Dim Month As Integer
Dim BranchID As Integer
TransType = IIf(optSales, 21, 22)
BranchID = Val(txtBranch)
Month = Val(txtMonth)
mYear = Val(txtYear)

    Dim mPosD As String
    Dim mServerD As String
     mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
     mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
     mServerD = ServerDb & ".dbo."
     
Dim mLastSerial As Double

Dim mLastSerialString As String
Dim intIal As String
Dim InBranch As String
Dim InitMonth As String
If BranchID < 10 Then
    InBranch = "0" + CStr(BranchID)
  

Else
    
    InBranch = CStr(BranchID)
End If

If Month < 10 Then

    InitMonth = "0" + CStr(Month)
Else
    InitMonth = CStr(Month)
End If



intIal = "101" + InBranch + "20" + InitMonth + "001"

        
s = " SELECT MAX(NoteSerial1) as aa"
s = s & " From Transactions"
s = s & " WHERE  Transaction_Type = " & TransType
s = s & "        AND MONTH(Transaction_Date) = " & Month
s = s & "                    AND YEAR(Transaction_Date) =" & mYear
s = s & "                    AND BranchId = " & BranchID
's = s & "                    AND SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50)), 1, 1) = " & BranchID
s = s & "                    AND ISNULL(NoteSerial1, '0') <> '0'"
s = s & "                    AND ISNULL(NoteSerial1, '') <> ''"
's = s & "                     AND BranchId = 9898"


s = s & "                     "

rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    mLastSerial = Val(rsDummy!aa & "")
    If mLastSerial = 0 Then
        mLastSerial = CDbl(intIal)
    End If
Else
    If mLastSerial = 0 Then
        mLastSerial = CLng(intIal)
    End If
End If
mLastSerialString = CStr(mLastSerial)

rsDummy.Close

    
    s = " SELECT *"
    s = s & " From Transactions"
    s = s & " WHERE  Transaction_Type = " & TransType
    s = s & "            AND YEAR(Transaction_Date) = " & mYear
    s = s & "            AND MONTH(Transactions.Transaction_Date) = " & Month
    s = s & "            AND BranchId = " & BranchID
    s = s & "            AND ISNULL(NoteSerial1, '0') IN (SELECT ISNULL(t.NoteSerial1, '0')"
    s = s & "                                             FROM   Transactions AS t"
    s = s & "                                             WHERE  t.Transaction_Type = " & TransType
    s = s & "                                                    AND YEAR(t.Transaction_Date) = " & mYear
    s = s & "                                                    AND MONTH(Transaction_Date) = " & Month
    s = s & "                                                    AND ISNULL(t.NoteSerial1, '0') IN (SELECT ISNULL(d.NoteSerial1, '0')"
    s = s & "                                                                                       FROM   Transactions d"
    s = s & "                                                                                       WHERE  d.Transaction_Type ="
    s = s & "                                                                                              " & TransType
    s = s & "                                                                                              AND MONTH(Transaction_Date) ="
    s = s & "                                                                                                  " & Month
    s = s & "                                                                                              AND BranchId = " & BranchID
    s = s & "                                                                                              AND YEAR(d.Transaction_Date) ="
    s = s & "                                                                                                  " & mYear & " )"
    s = s & "                                                    AND BranchId = " & BranchID
    s = s & "                                             Group By"
    s = s & "                                                    NoteSerial1"
    s = s & "                                             HAVING (COUNT(*) > 1))"




    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If rsDummy.EOF Then
        rsDummy.Close
        s = "Select *"
        s = s & " From Transactions"
        s = s & " WHERE  Transaction_Type = " & TransType
        s = s & "            AND YEAR(Transaction_Date) = " & mYear
        s = s & "            AND MONTH(Transactions.Transaction_Date) = " & Month
        s = s & "            AND BranchId = " & BranchID
        s = s & "            AND (IsNull(NoteSerial1,'') = '' Or IsNull(NoteSerial1,'0') = '0')"
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    End If
    Do While Not rsDummy.EOF
         rsDummy!NoteSerial1 = mLastSerial
         rsDummy!Account2 = "Test2"
         rsDummy.Update
         If Val(rsDummy!NoteId & "") <> 0 Then
            s = "UPDATE Notes Set NoteSerial1 = " & mLastSerial & " Where NoteId = " & Val(rsDummy!NoteId)
            Cn.Execute s
        End If
    
     
        mLastSerial = mLastSerial + 1
        
        rsDummy.MoveNext
    Loop

   frmPopup.ShowMessage " „ ÷»ÿ «·”Ì—Ì«·"
 
End Sub

Private Sub Command7_Click()
'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   Dim mWhere As String
   Dim mWhere2 As String
   Dim mWhere3 As String
   Dim mWhere4 As String
   
   
   Dim mWhere00 As String
   Dim mWhere22 As String
   Dim mWhere33 As String
   Dim mWhere44 As String
   'If txtFromDate.Value = txtToDate.Value Then
    mWhere = " NoteDate >= '" & SQLDate(txtFromDate.Value, False) & "' and NoteDate <= '" & SQLDate(txtToDate.Value, False) & "'"
    
    
    mWhere2 = " Transaction_Date>= '" & SQLDate(txtFromDate.Value, False) & "' and Transaction_Date<= '" & SQLDate(txtToDate.Value, False) & "'"
    mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."

    mWhere3 = " Notes_ID IN (SELECT DD.NoteId FROM " & mPosD & "Notes DD Where " & mWhere & ")"
    mWhere4 = " Transaction_ID IN (SELECT DD.Transaction_ID FROM " & mPosD & "Transactions DD Where " & mWhere2 & ")"

    
    
    mWhere00 = " TT.NoteDate >= '" & SQLDate(txtFromDate.Value, False) & "' and TT.NoteDate <= '" & SQLDate(txtToDate.Value, False) & "'"
    
    
    mWhere22 = " TT.Transaction_Date>= '" & SQLDate(txtFromDate.Value, False) & "' and TT.Transaction_Date<= '" & SQLDate(txtToDate.Value, False) & "'"
    
   

    'End If
    
    
    
  
  Cn.Execute "Delete notes_all where  " & mWhere
  Cn.Execute "Delete notes where " & mWhere
  Cn.Execute "Delete Transactions where " & mWhere2
   

    s = " update [" & POSlServer & "]." & POSDb & ".DBo.DOUBLE_ENTREY_VOUCHERS  set Double_Entry_Vouchers_ID  = Double_Entry_Vouchers_ID * -1"
    s = s & " where  RecordDate>= '" & SQLDate(txtFromDate.Value, False) & "' and RecordDate<= '" & SQLDate(txtToDate.Value, False) & "'"
    s = s & " and Double_Entry_Vouchers_ID > 0"
    
    POSConnection.Execute s

   
   
  UpdateFilesFromPos POSlServer, POSDb, "notes_all", "NoteID", mWhere00, mPosD
  UpdateFilesFromPos POSlServer, POSDb, "Transactions", "Transaction_ID", mWhere22, mPosD
  UpdateFilesFromPos POSlServer, POSDb, "notes", "NoteID", mWhere00, mPosD
  
  UpdateFilesFromPos POSlServer, POSDb, "Transaction_Details", "Transaction_ID", mWhere4, mPosD
  UpdateFilesFromPos POSlServer, POSDb, "TransactionValueAdded", "Transaction_ID", mWhere4, mPosD
  UpdateFilesFromPos POSlServer, POSDb, "DOUBLE_ENTREY_VOUCHERS", "DEV_ID_Line_No1", mWhere3, mPosD
  
  
  
  
   
  
   
End Sub

Private Sub Command8_Click()



On Error GoTo EE:
'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   
Dim mPosD  As String
Dim mServerD As String
mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."



mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
'mServerD = ServerDb & ".dbo."

'
's = " Update " & mPosD & "TblCustemers Set BranchId = 8 ,code = '003-' + Fullcode,Fullcode= '003-' + Fullcode"
's = s & " Where  CusName Not In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc  ) Or CusName"
's = s & " In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc where tc.CusID <> TblCustemers.CusID )"
'
'Cn.Execute s
'
's = " Update " & mPosD & "TblCustemers Set BranchId = 8 ,Fullcode= '003-' + Code"
's = s & " Where  CusName Not In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc  ) Or CusName"
's = s & " In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc where tc.CusID <> TblCustemers.CusID )"
'
'Cn.Execute s


Dim mWhere As String

mWhere = "   CusName Not In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc  ) Or CusName"
mWhere = mWhere & " In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc where tc.CusID <> TblCustemers.CusID )"




s = " Select * from  " & mPosD & "TblCustemers "
s = s & " Where  CusName Not In (Select Tc.CusName from  " & mServerD & "TblCustemers Tc  ) "
'Or Code "
's = s & " Not In (Select Tc.Code from  " & mServerD & "TblCustemers Tc  )"

Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
Dim rsDummy2 As New ADODB.Recordset

Dim mMaxId As Long
s = " Select Max(cusId) as MaxID from  " & mServerD & "TblCustemers "
mMaxId = 1
rsDummy2.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy2.EOF Then
    mMaxId = Val(rsDummy2!MaxId & "") + 1
End If
rsDummy2.Close
s = " Select * from  TblCustemers where cusId = -5"
rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
Do While Not rsDummy.EOF
    rsDummy2.AddNew
    rsDummy2!Code = rsDummy!Code & ""
    rsDummy2!CusName = rsDummy!CusName & ""
    rsDummy2!CusNamee = rsDummy!CusNamee & ""
    rsDummy2!FullCode = rsDummy!FullCode & ""
    rsDummy2!Cus_mobile = rsDummy!Cus_mobile & ""
    rsDummy2!Cus_Phone = rsDummy!Cus_Phone & ""
    rsDummy2!Remark = rsDummy!Remark & ""
    rsDummy2!Address = rsDummy!Address & ""
    rsDummy2!VATNO = rsDummy!VATNO & ""
    rsDummy2!BranchID = Val(rsDummy!BranchID & "")
    rsDummy2!CreditLimit = Val(rsDummy!CreditLimit & "")
    rsDummy2!OpenBalanceType = Val(rsDummy!OpenBalanceType & "")
    rsDummy2!CreditLimit = Val(rsDummy!CreditLimit & "")
    rsDummy2!Type = Val(rsDummy!Type & "")
    rsDummy2!cusID = mMaxId
  
    mMaxId = mMaxId + 1

    
    rsDummy2.Update


rsDummy.MoveNext
Loop

mServerD = ServerDb
Dim mPOSlServer As String
mPOSlServer = POSlServer.Text
'UpdateFilesFromPos ServerDb, POSDb, "TblCustemers", "CusId", mWhere, mPOSlServer
  
 Exit Sub
EE:
frmPopup.ShowMessage "BasicData"

End Sub

Private Sub Command9_Click()
On Error GoTo ErrTrap
Dim StrSQL As String
'On Error GoTo ErrTrap
If POSlServer.Text = "" Then
    MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
    Exit Sub
End If


If ConnectionFirst = False Then
Exit Sub
End If
Dim X As Date
Dim mTimeStart As String
'ServerDb = DestinationServer
 'POSDb = TxtServerDataBaseName.Text
    lblWait.Visible = True
  ' Command2_Click
    Dim rs As New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    JLCodeBasedOnBranch = IIf(rs("JLCodeBasedOnBranch").Value = 0 Or IsNull(rs("JLCodeBasedOnBranch").Value), False, True)
    StoreDigit = IIf(IsNull(rs("StoreDigit").Value), 1, (rs("StoreDigit").Value))
    BranchDigit = IIf(IsNull(rs("BranchDigit").Value), 1, (rs("BranchDigit").Value))
    

    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If SysSQLServerType = 1 Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
'            "Persist Security Info=False;Initial Catalog=" & POSDb & _
'            ";Data Source=" & POSlServer & ";Port=1433"
'
        .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
        ElseIf SysSQLServerType = 2 Then
             If SysSQLServerTypeTechnical = "0" Then
             .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                "Persist Security Info=False;Initial Catalog=" & POSDb & _
                ";Data Source=" & POSlServer & ";Port=1433"
              Else
                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer 'SysSQLServerName
            End If
        End If
       '   Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Adnan;Data Source=WAELPC\SQLEXPRESS;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=WAELPC;Use Encryption for Data=False;Tag with column collation when possible=False;

        .Open
    End With


GoTo Transactions

Transactions:
Dim SessionCode As String
Dim mMaxNo As Long
Dim ss As String
Dim rsDummyMax As New ADODB.Recordset
 Dim BeginTrans As Boolean
Dim isFoundData As Boolean

'ss = "Select Max(SessionCode ) MaxNo from TblOffline"
'rsDummyMax.Open ss, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'If rsDummyMax.EOF Then
'    mMaxNo = Val(rsDummyMax!MaxNo & "") + 1
    
'End If

SessionCode = CStr(Now) '& mMaxNo


'////////////////////////////////////////copy Sales Transactions

    Dim Rs3 As ADODB.Recordset
    Dim rsDouble_Entry As ADODB.Recordset
    
    Set Rs3 = New ADODB.Recordset
    
    
    Dim sql As String
    Dim mytext As String
    
   
   ' sql = " select * from Transactions    WHERE  Copied is null And " & GetQuery
    sql = " select * from Transactions    WHERE  Copied is null And POSBillType = 1 and " & GetQuery
    sql = " select * from Transactions    WHERE   Copied is null  And " & GetQuery
    
'    Dim tempString As String
'    Dim i As Integer
'    tempString = "0"
'    For i = 0 To Me.SelectedTransTypeList.ListCount - 1
'        tempString = tempString & "," & Me.SelectedTransTypeList.ItemData(i)
'    Next i
'    GetTransIds = tempString
    
    
    
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    mTimeStart = Now
    txtStartTime = mTimeStart
    Text3 = sql
    Dim FromTransaction_ID As Double
    Dim FromBranchID As Integer
    Dim FromTransaction_Date As Date
    Dim last_changed As Date
    
    Dim FromNots As String
    Dim FromNots2 As String
    Dim fromTransaction_Serial As String
    Dim FromNoteseial1 As String
    Dim FromTransaction_Type As Integer
    
    Dim BranchID As Integer
    Dim Transaction_ID As Double
    Dim Transaction_Type As Integer
    Dim Transaction_Date As Date
    Dim Transaction_Serial  As String
    Dim Nots As String
    Dim Nots2 As String
    Dim mTransaction_NetValue As Double
    Dim DepandToConv As Boolean
    Dim TypeInvoice As Integer
'eee
    'Dim Transaction_Type As Integer
    Dim FromNoteId As Double
   
    
   
 'sales
    If chkSales.Value = vbChecked Or chkSalesReturn.Value = vbChecked Or chkPurchase.Value = vbChecked Or chkPurchaseReturn.Value = vbChecked Or chkSalesOffers.Value = vbChecked Then
        If Rs3.RecordCount > 0 Then
'            Set cProgress = New ClsProgress
'            BolFrmLoaded = True
'            cProgress.ProgressType = Waiting
'            cProgress.StartProgress

'            Do While Rs3.State = adStateExecuting
'                DoEvents
'            Loop
            
'            If BolFrmLoaded = True Then
'                cProgress.StopProgess
'                Set cProgress = Nothing
'            End If

Dim rsCus As New ADODB.Recordset
Dim rsCus2 As New ADODB.Recordset

                Cn.BeginTrans
                BeginTrans = True
                
               ' MsgBox Rs3.RecordCount
                
                For i = 1 To Rs3.RecordCount
                    
                    
                    FromTransaction_Type = IIf(IsNull(Rs3("Transaction_Type").Value), 0, Rs3("Transaction_Type").Value)
                    FromTransaction_ID = IIf(IsNull(Rs3("Transaction_ID").Value), 0, Rs3("Transaction_ID").Value)
                    
                    Dim issueNoteid As String
                    Dim issuenoteserial As String
                    Dim issuenoteserial1 As String
                    Dim FromEmp_ID As Double
                    
                    Dim FromStoreID As Double
                    Dim FromCusID As Double
                    Dim FromBoxid As Double
                    Dim PayMentType As Integer
                    Dim BillBasedOn
                    'Dim BillBasedOn As Integer
                    Dim VATYou As Double
                    Dim VAT As Double
                    Dim FromUserID As Double
                    Dim POSBillType As Double
                    
                    Dim Trans_DiscountType As Integer
                    Dim Trans_Discount As Double
                    Dim TaxValue As Double
                    Dim order_no As String
                    Dim SaleType As Integer
                    Dim CashCustomerName As String
                    Dim TaxAddValue As Double
                    Dim CashCustomerPhone As String
                    Dim NetValue As Double
                    
                    Dim CarTypeID As Long, PlateNo As String, OilsTypesID As Long, YearFact As Long, Shaseh As String, CarMeter As String, FixesAssetsID As Long, ColorID2 As Integer, KM As Double, Chasee As String, PPointID As Long _
                    , Phone2 As String, SupplerID As Integer, Ser As Long, CarCurrentValue As Double, CarPrevValue As Double, CarEnginoil As Double, CarGearOil As Double, CarOilChangeDate As Date
                    
                    
                    
Dim PumpId As Long, PrevQty As Double, PrintName As String, Cash As Double, Mada As Double, Visa As Double, Deferred As Double, AmountH As Double, AmountHComm As Double, DetailsPump As String, Account_CodeComm As String, _
    Account_Code As String, IsOther As Boolean
                    
                    FromUserID = IIf(IsNull(Rs3("UserID").Value), 0, Rs3("UserID").Value)
                    FromEmp_ID = IIf(IsNull(Rs3("Emp_ID").Value), 0, Rs3("Emp_ID").Value)
                    FromStoreID = IIf(IsNull(Rs3("storeID").Value), 0, Rs3("storeID").Value)
                    CarTypeID = Val(Rs3!CarTypeID & "")
                    PlateNo = Trim(Rs3!PlateNo & "")
                    OilsTypesID = Val(Rs3!OilsTypesID & "")
                    YearFact = Val(Rs3!YearFact & "")
                    Shaseh = Trim(Rs3!Shaseh & "")
                    CarMeter = Trim(Rs3!CarMeter & "")
                    FixesAssetsID = Val(Rs3!FixesAssetsID & "")
                    ColorID2 = Val(Rs3!ColorID2 & "")
                    Chasee = Trim(Rs3!Chasee & "")
                    KM = Val(Rs3!KM & "")
                    PPointID = Val(Rs3!KM & "")
                    Phone2 = Trim(Rs3!Phone2 & "")
                    SupplerID = Val(Rs3!SupplerID & "")
                    Ser = Val(Rs3!Ser & "")
                    CarCurrentValue = Val(Rs3!CarCurrentValue & "")
                    CarPrevValue = Val(Rs3!CarPrevValue & "")
                    CarEnginoil = Val(Rs3!CarEnginoil & "")
                    CarGearOil = Val(Rs3!CarGearOil & "")
                    CarOilChangeDate = IIf(IsNull(Rs3("CarOilChangeDate").Value), Date, Rs3("CarOilChangeDate").Value)
                    
                    FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                    If FromTransaction_Type = 42 Then
                        s = "Select Code,CusName from TblCustemers where cusId =  " & FromCusID
                        Set rsCus = New ADODB.Recordset
                        rsCus.Open s, POSConnection, adOpenStatic, adLockReadOnly
                        If Not rsCus.EOF Then
                            s = "Select cusId,Code,CusName from TblCustemers where Code =  N'" & Trim(rsCus!Code & "") & "' and CusName =  N'" & Trim(rsCus!CusName & "") & "'"
                            Set rsCus2 = New ADODB.Recordset
                            rsCus2.Open s, Cn, adOpenStatic, adLockReadOnly
                            If Not rsCus2.EOF Then
                                FromCusID = Val(rsCus2!cusID & "")
                            End If
                            
                        End If
                        
                        
                    Else
                        FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                    End If
                    
                    FromBoxid = IIf(IsNull(Rs3("Boxid").Value), 0, Rs3("Boxid").Value)
                    POSBillType = IIf(IsNull(Rs3("POSBillType").Value), 0, Rs3("POSBillType").Value)
                    FromUserID = Val(Rs3!userID & "")
                    mTransaction_NetValue = Val(Rs3!Transaction_NetValue & "")
                    FromPaymentType = IIf(IsNull(Rs3("PaymentType").Value), 0, Rs3("PaymentType").Value)
                    FromBillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                    FromVATYou = IIf(IsNull(Rs3("VATYou").Value), 0, Rs3("VATYou").Value)
                    FromVAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
                    TypeInvoice = IIf(IsNull(Rs3("TypeInvoice").Value), 0, Rs3("TypeInvoice").Value)
                    '
                    BillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                    DepandToConv = True
                    Trans_DiscountType = IIf(IsNull(Rs3("Trans_DiscountType").Value), 0, Rs3("Trans_DiscountType").Value)
                    Trans_Discount = IIf(IsNull(Rs3("Trans_Discount").Value), 0, Rs3("Trans_Discount").Value)
                    TaxValue = IIf(IsNull(Rs3("TaxValue").Value), 0, Rs3("TaxValue").Value)
                    SaleType = IIf(IsNull(Rs3("SaleType").Value), 0, Rs3("SaleType").Value)
                    TaxAddValue = IIf(IsNull(Rs3("TaxAddValue").Value), 0, Rs3("TaxAddValue").Value)
                    
                    CashCustomerName = IIf(IsNull(Rs3("CashCustomerName").Value), "", Rs3("CashCustomerName").Value)
                    CashCustomerPhone = IIf(IsNull(Rs3("CashCustomerPhone").Value), "", Rs3("CashCustomerPhone").Value)
                    order_no = IIf(IsNull(Rs3("order_no").Value), "", Rs3("order_no").Value)
     
                    
                    FromBranchID = IIf(IsNull(Rs3("BranchID").Value), 0, Rs3("BranchID").Value)
                    fromTransaction_Serial = IIf(IsNull(Rs3("Transaction_Serial").Value), 0, Rs3("Transaction_Serial").Value)
                    
                    
                    FromNoteSerial1 = IIf(IsNull(Rs3("Noteserial1").Value), 0, Rs3("Noteserial1").Value)
                    FromNoteSerial = IIf(IsNull(Rs3("Noteserial").Value), 0, Rs3("Noteserial").Value)
                    'FromNoteId = IIf(IsNull(Rs3("NoteId").Value), 0, Rs3("NoteId").Value) ' —ﬁ„ ﬁÌœ «·›« Ê—…
                    
                    FromNots = IIf(IsNull(Rs3("Nots").Value), 0, Rs3("Nots").Value) '—ﬁ„ ”‰œ «·’—›
                    If FromTransaction_Type <> 42 Then
                        GetIssueData CDbl(Val(FromNots)), issueNoteid, issuenoteserial, issuenoteserial1
                    End If
                    FromNots2 = IIf(IsNull(Rs3("Nots2").Value), 0, Rs3("Nots2").Value)
                    FromTransaction_Date = IIf(IsNull(Rs3("Transaction_Date").Value), 0, Rs3("Transaction_Date").Value)
                    last_changed = IIf(IsNull(Rs3("last_changed").Value), Date, Rs3("last_changed").Value)
                    NetValue = IIf(IsNull(Rs3("NetValue").Value), 0, Rs3("NetValue").Value)
                    Transaction_Date = FromTransaction_Date
                    Transaction_Type = FromTransaction_Type
                    
                    Select Case Transaction_Type
                    Case 21
                        mNoteType = 170
                        mSanadNo = 7
                        CountSales = CountSales + 1
                    Case 19
                        mNoteType = 180
                        mSanadNo = 10
                        'CountSales =CountSales +1
                    Case 20
                        mNoteType = 160
                        mSanadNo = 9
                    
                    Case 5
                        mNoteType = 230
                        mSanadNo = 15
                        CountPurchaseReturn = CountPurchaseReturn + 1
                    Case 22
                        mNoteType = 150
                        mSanadNo = 6
                        CountPurchase = CountPurchase + 1
                    Case 9
                        mNoteType = 220
                        mSanadNo = 14
                        CountSalesReturn = CountSalesReturn + 1
                    Case 42
                        mNoteType = 0
                        mSanadNo = 42
                        CountSalesOfeers = CountSalesOfeers + 1
                    
                    End Select
                    isFoundData = True
                    
                    BranchID = FromBranchID
                    Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
                    'Transaction_Serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=" & Transaction_Type & ""))
                    'NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , Transaction_Type, , , , , , FromUserID)
                    NoteSerial1 = Rs3!NoteSerial1 & ""
                    If POSBillType = 0 Then
                        NoteSerial = Notes_coding(BranchID, FromTransaction_Date)
                        NoteId = CStr(new_id("Notes", "NoteID", "", True))
                    End If
                    TransactionComment = " ”‰œ  ÕÊÌ· „‰ﬁÊ·… „‰ ﬁ«⁄œ…  " & POSname.Text & "   "
                    TransactionComment = TransactionComment & "  —ﬁ„ «·›« Ê—…  «·«’·Ì…" & FromNoteSerial1
                 '" & ServerDb & "
                 '   MsgBox TransactionComment
     
                    '*****************************************************
                    '*****************************************************
                    If Trim(NoteSerial) = "" Then NoteSerial = "0"
                    If Val(NoteId) = 0 Then NoteId = 0
                    'ÂÌœ— «·›« Ê—…
                    '*****************************************************************************************
                   

                    
                    
                    sql = " INSERT INTO  [" & POSDb & "].[dbo].[Transactions]  (    "
                    sql = sql & "  Transaction_ID,Transaction_Date,TypeInvoice,"
                    sql = sql & "   Transaction_Serial ,"
                    sql = sql & "   Transaction_Type, "
                    sql = sql & "  PaymentType,"
                    sql = sql & "   CusID, StoreID, "
                    sql = sql & "  UserID, Emp_ID, "
                    sql = sql & "  BranchId, BoxID , "
                    sql = sql & "  BillBasedOn, VAT, "
                    sql = sql & "  VATYou, NoteSerial,"
                    sql = sql & "  NoteSerial1,NoteId,"
                    sql = sql & "  Copied,TransactionComment,"
                    sql = sql & " SessionCode,POSBillType, "
                    sql = sql & "  OldNoteserial1,OldNoteserial,"
                    sql = sql & " OldNoteId,OldTransaction_ID,"
                    
                    sql = sql & " Trans_DiscountType  ,"
                    sql = sql & " Trans_Discount   ,"
                    sql = sql & "TaxValue  ,"
                    sql = sql & " order_no  ,"
                    sql = sql & " SaleType  ,"
                    sql = sql & " CashCustomerName  ,"
                    sql = sql & "TaxAddValue  ,"
                    sql = sql & "CashCustomerPhone,last_changed ,NetValue,Transaction_NetValue,DepandToConv ,"
                    
                    
                    
                    sql = sql & "CarTypeID,"
                    sql = sql & "PlateNo,"
                    sql = sql & "OilsTypesID,"
                    sql = sql & "YearFact,"
                    sql = sql & "Shaseh,"
                    sql = sql & "CarMeter,"
                    sql = sql & "FixesAssetsID,"
                    sql = sql & "ColorID2,"
                    sql = sql & "KM,"
                    sql = sql & "Chasee,"
                    sql = sql & "PPointID,"
                    sql = sql & "Phone2,"
                    sql = sql & "SupplerID,"
                    sql = sql & "Ser,"
                    sql = sql & "CarCurrentValue,"
                    sql = sql & "CarPrevValue,"
                    sql = sql & "CarEnginoil,"
                    sql = sql & "CarGearOil,"
                    sql = sql & "CarOilChangeDate                    )"
                    
                    
                    
                    
                    sql = sql & "   values (" & Transaction_ID & "," & SQLDate(Transaction_Date, True) & "," & TypeInvoice & ","
                    sql = sql & FromNoteSerial1 & ","
                    sql = sql & Transaction_Type & ","
                    sql = sql & FromPaymentType & ","
                    sql = sql & FromCusID & "," & FromStoreID & ","
                    sql = sql & FromUserID & "," & FromEmp_ID & ","
                    sql = sql & FromBranchID & "," & FromBoxid
                    sql = sql & "," & BillBasedOn & "," & FromVAT & ","
                    sql = sql & FromVATYou & ","
                    sql = sql & NoteSerial & ",'" & FromNoteSerial1 & "'," & NoteId & ",1,'"
                    sql = sql & TransactionComment & "','" & SessionCode & "', " & POSBillType & " , '" & FromNoteSerial1 & "' , "
                    sql = sql & "'" & FromNoteSerial & "' , " & FromNoteId & " , " & FromTransaction_ID & " ,"
                    sql = sql & Trans_DiscountType & " ,"
                    sql = sql & Trans_Discount & " ,"
                    sql = sql & TaxValue & " ,"
                    sql = sql & "'" & order_no & "' ,"
                    sql = sql & SaleType & " ,"
                    sql = sql & "'" & CashCustomerName & "' ,"
                    sql = sql & TaxAddValue & " ,"
                    sql = sql & "'" & CashCustomerPhone & "' ," & SQLDate(last_changed, True) & "," & NetValue & "," & mTransaction_NetValue & "," & IIf(DepandToConv, 1, 0) & ", "
                    
                    
                    sql = sql & CarTypeID & ","
                    sql = sql & "'" & PlateNo & "',"
                    sql = sql & OilsTypesID & ","
                    sql = sql & YearFact & ","
                    sql = sql & "'" & Shaseh & "',"
                    sql = sql & "'" & CarMeter & "',"
                    sql = sql & FixesAssetsID & ","
                    sql = sql & ColorID2 & ","
                    sql = sql & KM & ","
                    sql = sql & "'" & Chasee & "',"
                    sql = sql & PPointID & ","
                    sql = sql & "'" & Phone2 & "',"
                    sql = sql & SupplerID & ","
                    sql = sql & Ser & ","
                    sql = sql & CarCurrentValue & ","
                    sql = sql & CarPrevValue & ","
                    sql = sql & CarEnginoil & ","
                    sql = sql & CarGearOil & ","
                    sql = sql & SQLDate(CarOilChangeDate, True) & "                       )"
                    
                    
       



                    
                    
                    '   fromTransaction_Serial
                    Text1.Text = sql
                   ' Exit Sub
                   Text4 = ""
                    POSConnection.Execute sql
                    Text4 = sql
                    
                    ' ›«’Ì· «·›« Ê—…
                    
                    
                    sql = " select * from Transaction_Details   where  Transaction_ID=" & FromTransaction_ID
                    Set rsDouble_Entry = New ADODB.Recordset
                    '
                    rsDouble_Entry.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Dim j As Double
                    For j = 1 To rsDouble_Entry.RecordCount
                        Item_ID = IIf(IsNull(rsDouble_Entry("Item_ID").Value), 0, rsDouble_Entry("Item_ID").Value)
                        ItemCase = IIf(IsNull(rsDouble_Entry("ItemCase").Value), 0, rsDouble_Entry("ItemCase").Value)
                        Quantity = IIf(IsNull(rsDouble_Entry("Quantity").Value), 0, rsDouble_Entry("Quantity").Value)
                        Price = IIf(IsNull(rsDouble_Entry("Price").Value), 0, rsDouble_Entry("Price").Value)
                        ItemDiscountType = IIf(IsNull(rsDouble_Entry("ItemDiscountType").Value), 0, rsDouble_Entry("ItemDiscountType").Value)
                        ItemDiscount = IIf(IsNull(rsDouble_Entry("ItemDiscount").Value), 0, rsDouble_Entry("ItemDiscount").Value)
                        ShowQty = IIf(IsNull(rsDouble_Entry("ShowQty").Value), 0, rsDouble_Entry("ShowQty").Value)
                        showPrice = IIf(IsNull(rsDouble_Entry("showPrice").Value), 0, rsDouble_Entry("showPrice").Value)
                        UnitID = IIf(IsNull(rsDouble_Entry("UnitId").Value), 0, rsDouble_Entry("UnitId").Value)
                        ColorID = IIf(IsNull(rsDouble_Entry("ColorID").Value), 0, rsDouble_Entry("ColorID").Value)
                        ItemSize = IIf(IsNull(rsDouble_Entry("ItemSize").Value), 0, rsDouble_Entry("ItemSize").Value)
                        ClassId = IIf(IsNull(rsDouble_Entry("ClassId").Value), 0, rsDouble_Entry("ClassId").Value)
                        mmVatyo = IIf(IsNull(rsDouble_Entry("Vatyo").Value), 0, rsDouble_Entry("Vatyo").Value)
                    
                        PumpId = Val(rsDouble_Entry!PumpId & "")
                        PrevQty = Val(rsDouble_Entry!PrevQty & "")
                        PrintName = Trim(rsDouble_Entry!PrintName & "")
                        Cash = Val(rsDouble_Entry!Cash & "")
                        Mada = Val(rsDouble_Entry!Mada & "")
                        Visa = Val(rsDouble_Entry!Visa & "")
                        Deferred = Val(rsDouble_Entry!Deferred & "")
                        AmountH = Val(rsDouble_Entry!AmountH & "")
                        AmountHComm = Val(rsDouble_Entry!AmountHComm & "")
                        DetailsPump = Trim(rsDouble_Entry!DetailsPump & "")
                        Account_CodeComm = Trim(rsDouble_Entry!Account_CodeComm & "")
                        Account_Code = Trim(rsDouble_Entry!Account_Code & "")
                        IsOther = Val(rsDouble_Entry!IsOther & "")
                        
                        sql = " INSERT INTO  [" & POSDb & "].[dbo].[Transaction_Details]  (    "
                        sql = sql & "  Transaction_ID,  Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice,UnitId , ColorID, ItemSize, ClassId,SessionCode,Vatyo,"
                        
                        sql = sql & "  PumpId,"
                        sql = sql & "  PrevQty,"
                        sql = sql & "  PrintName,"
                        sql = sql & "  Cash,"
                        sql = sql & "  Mada,"
                        sql = sql & "  Visa,"
                        sql = sql & "  Deferred,"
                        sql = sql & "  AmountH,"
                        sql = sql & "  AmountHComm,"
                        sql = sql & "  DetailsPump,"
                        sql = sql & "  Account_CodeComm,"
                        sql = sql & "  Account_Code,"
                        sql = sql & "  IsOther"
                        
                        sql = sql & "  )"
                        sql = sql & "   values (" & Transaction_ID & "," & Item_ID & ", " & ItemCase & "," & Quantity & "," & Price & "," & ItemDiscountType & "," & ItemDiscount & "," & ShowQty & "," & showPrice
                        sql = sql & "," & UnitID & "," & ColorID & "," & ItemSize & "," & ClassId & "" & ",'" & SessionCode & "'," & mmVatyo & ","
                        
                        sql = sql & PumpId & ","
                        sql = sql & PrevQty & ","
                        sql = sql & "'" & PrintName & "',"
                        sql = sql & Cash & ","
                        sql = sql & Mada & ","
                        sql = sql & Visa & ","
                        sql = sql & Deferred & ","
                        sql = sql & AmountH & ","
                        sql = sql & AmountHComm & ","
                        sql = sql & "'" & DetailsPump & "',"
                        sql = sql & "'" & Account_CodeComm & "',"
                        sql = sql & "'" & Account_Code & "',"
                        sql = sql & IIf(IsOther, 1, 0) & ")"
                        
                        POSConnection.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j
              '      MsgBox "3"
    '*********************** ›«’Ì· «·›«  ********************************************************
                    sql = " select * from TransactionValueAdded   where  Transaction_ID=" & FromTransaction_ID
                    Set rsDouble_Entry = New ADODB.Recordset
                    '
                    rsDouble_Entry.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    'Dim j As Double
                    For j = 1 To rsDouble_Entry.RecordCount
                        ItemID = IIf(IsNull(rsDouble_Entry("ItemID").Value), 0, rsDouble_Entry("ItemID").Value)
                        Vatyo = IIf(IsNull(rsDouble_Entry("Vatyo").Value), 0, rsDouble_Entry("Vatyo").Value)
                        VAT = IIf(IsNull(rsDouble_Entry("Vat").Value), 0, rsDouble_Entry("Vat").Value)
                        Valu = IIf(IsNull(rsDouble_Entry("Valu").Value), 0, rsDouble_Entry("Valu").Value)
                        selectd = IIf(IsNull(rsDouble_Entry("selectd").Value), 0, rsDouble_Entry("selectd").Value)
                        
                        
                        sql = " INSERT INTO  [" & POSDb & "].[dbo].[TransactionValueAdded]  (    "
                        sql = sql & "  Transaction_ID,  ItemID, Vatyo, VAT, Valu, selectd,Transaction_Type,SessionCode)"
                        sql = sql & "   values (" & Transaction_ID & "," & ItemID & ", " & Vatyo & "," & VAT & "," & Valu & "," & selectd & "," & Transaction_Type & " ,'" & SessionCode & "' )"
                        
                        
                        POSConnection.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j
  
    '*********************** ›«’Ì· «·›«  ********************************************************
                    
     
    '*********************** ›«’Ì· «·‘»ﬂ… ********************************************************
                    If Transaction_Type = 21 Or Transaction_Type = 9 Then
                        sql = " select * from TblTransactionPayments   where  Transaction_ID=" & FromTransaction_ID
                        Set rsDouble_Entry = New ADODB.Recordset
                        '
                        Dim Recorddate As Date
                        rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                        'Dim j As Double
                        For j = 1 To rsDouble_Entry.RecordCount
                            BoxID = IIf(IsNull(rsDouble_Entry("boxid").Value), 0, rsDouble_Entry("boxid").Value)
                            Recorddate = IIf(IsNull(rsDouble_Entry("Recorddate").Value), 0, rsDouble_Entry("Recorddate").Value)
                            PointID = IIf(IsNull(rsDouble_Entry("PointID").Value), 0, rsDouble_Entry("PointID").Value)
                            CurrentCashireID = IIf(IsNull(rsDouble_Entry("CurrentCashireID").Value), 0, rsDouble_Entry("CurrentCashireID").Value)
                            PaymentID = IIf(IsNull(rsDouble_Entry("PaymentID").Value), 0, rsDouble_Entry("PaymentID").Value)
                            Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                            CardNo = IIf(IsNull(rsDouble_Entry("CardNo").Value), 0, rsDouble_Entry("CardNo").Value)
                            Effect = IIf(IsNull(rsDouble_Entry("Effect").Value), 0, rsDouble_Entry("Effect").Value)
                            
                            
                            sql = " INSERT INTO  [" & ServerDb & "].[dbo].[TblTransactionPayments]  (    "
                            sql = sql & "  Transaction_ID,  boxid, Recorddate, PointID, CurrentCashireID, PaymentID,Value,CardNo,Effect,SessionCode)"
                            sql = sql & "   values (" & Transaction_ID & "," & BoxID & ", " & SQLDate(Recorddate, True) & "," & PointID & "," & CurrentCashireID & "," & PaymentID & "," & Value & ",'" & CardNo & "'," & Effect & ",'" & SessionCode & "')"
                            
                            
                            Cn.Execute sql
                            rsDouble_Entry.MoveNext
                        Next j
                 '       MsgBox "5"
                    End If
'                    MsgBox "3"
    '*********************** ›«’Ì·  «·‘»ﬂ… ********************************************************
      
      
             
                'ﬁÌœ «·›« Ê—…
                 
                If POSBillType = 0 And Transaction_Type <> 42 Then
                 
                 
                    sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,Transaction_ID,UserID,SessionCode)"
                    sql = sql & " values( " & NoteId & ", " & SQLDate(Transaction_Date, True) & " ,  " & mNoteType & ", '" & NoteSerial & "', '" & NoteSerial1 & "'," & BranchID & "," & Transaction_ID & ",1,'" & SessionCode & "')"
                    Cn.Execute sql
                    DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                    
                    
                    'Dim rsDouble_Entry As ADODB.Recordset
                    Set rsDouble_Entry = New ADODB.Recordset
                    sql = " select * from DOUBLE_ENTREY_VOUCHERS   where   Notes_ID=" & FromNoteId
                    rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    Dim w As Double
                    For w = 1 To rsDouble_Entry.RecordCount
                        Account_Code = IIf(IsNull(rsDouble_Entry("Account_Code").Value), 0, rsDouble_Entry("Account_Code").Value)
                        Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                        Credit_Or_Debit = IIf(IsNull(rsDouble_Entry("Credit_Or_Debit").Value), 0, rsDouble_Entry("Credit_Or_Debit").Value)
                        Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                        Double_Entry_Vouchers_Description = IIf(IsNull(rsDouble_Entry("Double_Entry_Vouchers_Description").Value), 0, rsDouble_Entry("Double_Entry_Vouchers_Description").Value) & Chr(13) & "  ”‰œ ’—› " & TransactionComment
                        'RecordDate = IIf(IsNull(rsDouble_Entry("RecordDate").Value), 0, rsDouble_Entry("RecordDate").Value)
                        DEV_ID_Line_No = IIf(IsNull(rsDouble_Entry("DEV_ID_Line_No").Value), 0, rsDouble_Entry("DEV_ID_Line_No").Value)
                        branch_id = IIf(IsNull(rsDouble_Entry("branch_id").Value), 0, rsDouble_Entry("branch_id").Value)
                        sql = "  INSERT INTO [" & ServerDb & "].[dbo].[DOUBLE_ENTREY_VOUCHERS]([Double_Entry_Vouchers_ID], [DEV_ID_Line_No], [Account_Code], [Value], [Credit_Or_Debit], [Double_Entry_Vouchers_Description], [RecordDate], [Notes_ID] ,branch_id,UserID,Transaction_ID,SessionCode) "
                        sql = sql & " values (  " & DEVID & ", " & DEV_ID_Line_No & ", '" & Account_Code & "', " & Value & ", " & Credit_Or_Debit & ", '" & Double_Entry_Vouchers_Description & "',  " & SQLDate(Transaction_Date, True) & ", " & NoteId & " ," & branch_id & "," & 1 & "," & Transaction_ID & ",'" & SessionCode & "')"
                        Cn.Execute sql
                        
                        
                        rsDouble_Entry.MoveNext
                    Next w
               '     MsgBox "6"
                    '*****************************************************************
                '**********************************************************
                 End If
    
      
              '     GetIssueData CDbl(FromNots), issueNoteid, issuenoteserial, issuenoteserial1
                Dim mTransType2 As Integer
                If Transaction_Type = 21 Or Transaction_Type = 5 Then
                    mTransType2 = 19
                Else
                    mTransType2 = 20
                End If
                If chkDontCopyIss.Value = vbUnchecked And Transaction_Type <> 42 Then
                    CopyIssueTtransaction Transaction_ID, CStr(NoteSerial1), CDbl(Val(FromNots)), CDbl(mTransType2), issuenoteserial, issuenoteserial1, SessionCode
                End If
             
                Rs3.MoveNext
             
                lblCount.Caption = Val(lblCount.Caption) + 1
            Next i
   '      MsgBox "7"
         
        End If
     
        Rs3.Close
      'Sql = Sql & "[" & POSDb & "].dbo.Transactions"
      '„‰⁄ «·‰ﬁ· „—… «Œ—Ì
      
    
            sql = "update   [" & POSDb & "].dbo.Transactions" & "  set  Copied =1,SessionCode = '" & SessionCode & "' "
      sql = sql & "  Where Copied Is Null And "
      sql = sql & GetQuery
      '& "  and dbo.Transactions.Transaction_Date ='" & SQLDate(dbRecordDate.Value, False) & "'"
      
     POSConnection.Execute sql
    ' MsgBox "8"
    
'      sql = "update   [" & POSDb & "].dbo.Transaction_Details" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE  Copied is null    "
'     POSConnection.Execute sql
     
End If
  

If chkRec.Value = vbChecked Then
 sql = " select * From Notes where NoteType=4   and   Copied is null " ' ”‰œ«  ﬁ»÷
 
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    
     If Rs3.RecordCount > 0 Then
      
            For i = 1 To Rs3.RecordCount
                CashingType = IIf(IsNull(Rs3("CashingType").Value), 0, Rs3("CashingType").Value)
                          
               FromNoteId = IIf(IsNull(Rs3("NoteID").Value), 0, Rs3("NoteID").Value)
               FromBranchID = IIf(IsNull(Rs3("Branch_no").Value), 0, Rs3("Branch_no").Value)
               
              
                FromNoteSerial1 = IIf(IsNull(Rs3("Noteserial1").Value), 0, Rs3("Noteserial1").Value)
                FromNoteSerial = IIf(IsNull(Rs3("Noteserial").Value), 0, Rs3("Noteserial").Value)
                BranchID = FromBranchID
                NoteDate = IIf(IsNull(Rs3("NoteDate").Value), 0, Rs3("NoteDate").Value)
                'notedate = FromTransaction_Date
                NoteSerial = Notes_coding(BranchID, CDate(NoteDate))
                DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                'Dim notedate1 As Date
              
                NoteSerial1 = Voucher_coding(CInt(BranchID), CDate(NoteDate), 2, 4, , , , , , , , FromUserID)
         
                NoteId = CStr(new_id("Notes", "NoteID", "", True))
                TransactionComment = " ”‰œ ﬁ»÷  „‰ﬁÊ· „‰ ﬁ«⁄œ…  " & POSDb & "   "
                TransactionComment = TransactionComment & "  —ﬁ„ «·”‰œ  «·«’·Ì" & FromNoteSerial1
                TransactionComment = TransactionComment & "  —ﬁ„ «·ﬁÌœ  «·«’·Ì" & FromNoteSerial
   
                EmpId = IIf(IsNull(Rs3("EmpId").Value), 0, Rs3("EmpId").Value)
                VAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
                person = IIf(IsNull(Rs3("person").Value), 0, Rs3("person").Value)
                NCashingType = IIf(IsNull(Rs3("NCashingType").Value), 0, Rs3("NCashingType").Value)
                Status = IIf(IsNull(Rs3("Status").Value), 0, Rs3("Status").Value)
                Note_Value = IIf(IsNull(Rs3("Note_Value").Value), 0, Rs3("Note_Value").Value)
                BankName = IIf(IsNull(Rs3("BankName").Value), 0, Rs3("BankName").Value)
                Remark = IIf(IsNull(Rs3("Remark").Value), 0, Rs3("Remark").Value)
                cusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                NoteCashingType = IIf(IsNull(Rs3("NoteCashingType").Value), 0, Rs3("NoteCashingType").Value)
                BoxID = IIf(IsNull(Rs3("BoxID").Value), "Null", Rs3("BoxID").Value)
                ChqueNum = IIf(IsNull(Rs3("ChqueNum").Value), 0, Rs3("ChqueNum").Value)
                DueDate = IIf(IsNull(Rs3("DueDate").Value), 0, Rs3("DueDate").Value)
                ChequeBoxID = IIf(IsNull(Rs3("ChequeBoxID").Value), 0, Rs3("ChequeBoxID").Value)
                BankID = IIf(IsNull(Rs3("BankID").Value), "Null", Rs3("BankID").Value)
                TotalNotesValue = IIf(IsNull(Rs3("TotalNotesValue").Value), 0, Rs3("TotalNotesValue").Value)
                
                sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,UserID,CashingType,EmpId,VAT"
                 sql = sql & ",NCashingType, Status,Note_Value,BankName,Remark,CusID,NoteCashingType,BoxID,ChqueNum,DueDate,ChequeBoxID,BankID,TotalNotesValue,copied,SessionCode )"
                 sql = sql & " values( " & NoteId & ", " & SQLDate(CDate(NoteDate), True) & " , 4, " & NoteSerial & ", " & NoteSerial1 & "," & BranchID & ",1," & CashingType & "," & EmpId & "," & VAT
                 sql = sql & "," & NCashingType & ", " & Status & "," & Note_Value & ",'" & BankName & "','" & Remark & "'," & cusID & "," & NoteCashingType & "," & BoxID & "," & ChqueNum & "," & SQLDate(CDate(Date), True) & "," & ChequeBoxID & "," & BankID & "," & TotalNotesValue & ",1,'" & SessionCode & "')"
                 Cn.Execute sql
                 DEVID = CStr(new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", ""))
                 
                 
                 'Dim rsDouble_Entry As ADODB.Recordset
                  Set rsDouble_Entry = New ADODB.Recordset
                     sql = " select * from DOUBLE_ENTREY_VOUCHERS   where   Notes_ID=" & FromNoteId
                   rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    'Dim w As Double
                    For w = 1 To rsDouble_Entry.RecordCount
                    Account_Code = IIf(IsNull(rsDouble_Entry("Account_Code").Value), 0, rsDouble_Entry("Account_Code").Value)
                    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                    Credit_Or_Debit = IIf(IsNull(rsDouble_Entry("Credit_Or_Debit").Value), 0, rsDouble_Entry("Credit_Or_Debit").Value)
                    Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                    Double_Entry_Vouchers_Description = IIf(IsNull(rsDouble_Entry("Double_Entry_Vouchers_Description").Value), 0, rsDouble_Entry("Double_Entry_Vouchers_Description").Value) & Chr(13) & "    " & TransactionComment
                    Recorddate = IIf(IsNull(rsDouble_Entry("RecordDate").Value), 0, rsDouble_Entry("RecordDate").Value)
                    DEV_ID_Line_No = IIf(IsNull(rsDouble_Entry("DEV_ID_Line_No").Value), 0, rsDouble_Entry("DEV_ID_Line_No").Value)
                    branch_id = IIf(IsNull(rsDouble_Entry("branch_id").Value), 0, rsDouble_Entry("branch_id").Value)
                    sql = "  INSERT INTO [" & ServerDb & "].[dbo].[DOUBLE_ENTREY_VOUCHERS]([Double_Entry_Vouchers_ID], [DEV_ID_Line_No], [Account_Code], [Value], [Credit_Or_Debit], [Double_Entry_Vouchers_Description], [RecordDate], [Notes_ID] ,branch_id,UserID ,SessionCode ) "
                    sql = sql & " values (  " & DEVID & ", " & DEV_ID_Line_No & ", '" & Account_Code & "', " & Value & ", " & Credit_Or_Debit & ", '" & Double_Entry_Vouchers_Description & "',  " & SQLDate(Recorddate, True) & ", " & NoteId & " ," & branch_id & ", 1,'" & SessionCode & "')"
                    Cn.Execute sql
                
                
                    rsDouble_Entry.MoveNext
                    Next w
                
                
                '*********************** ›«’Ì· «·‘Ìﬂ… ··ﬁ»÷ ********************************************************
                    sql = " select * from TblMultuPayment   where  NoteID=" & FromNoteId
                    Set rsDouble_Entry = New ADODB.Recordset
                    '
                     
                    rsDouble_Entry.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
                    'Dim j As Double
                    For j = 1 To rsDouble_Entry.RecordCount
                '     NoteId = IIf(IsNull(rsDouble_Entry("NoteId").Value), 0, rsDouble_Entry("NoteId").Value)
                        PaymentID = IIf(IsNull(rsDouble_Entry("PaymentID").Value), 0, rsDouble_Entry("PaymentID").Value)
                        Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                        CardNo = IIf(IsNull(rsDouble_Entry("CardNo").Value), " ", rsDouble_Entry("CardNo").Value)
                        maxvalue = IIf(IsNull(rsDouble_Entry("maxvalue").Value), 0, rsDouble_Entry("maxvalue").Value)
                        sql = " INSERT INTO  [" & ServerDb & "].[dbo].[TblMultuPayment]  (    "
                        sql = sql & "  NoteId,   PaymentID, Value, CardNo, maxvalue ,SessionCode )"
                        sql = sql & "   values (" & NoteId & ", " & PaymentID & "," & Value & ",'" & CardNo & "'," & maxvalue & ",'" & SessionCode & "')"
                        
                        
                        Cn.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j

'*********************** ··ﬁ»÷  ›«’Ì·  «·‘»ﬂ… ********************************************************
  
  
 
            Next i
            
      End If
        sql = "update   [" & POSDb & "].dbo.Notes" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE   NoteType=4   and   Copied is null  "
        sql = sql & " and dbo.Notes.NoteDate ='" & SQLDate(dbRecordDate.Value, False) & "'"
        POSConnection.Execute sql
     '   MsgBox "9"
        
'        sql = "update   [" & POSDb & "].dbo.DOUBLE_ENTREY_VOUCHERS" & "  set  Copied =1,SessionCode = '" & SessionCode & "' WHERE  Copied is null   "
'        POSConnection.Execute sql


 End If
If isFoundData Then
     Dim rsOffline As New ADODB.Recordset
    Dim mEndTime22 As String
    mEndTime22 = Now
    s = "Select * from TblOffline where 1 = -1"
    rsOffline.Open s, Cn, adOpenKeyset, adLockOptimistic
    'MsgBox s
    rsOffline.AddNew
    'MsgBox s & "Save"
    'rsOffline!Id = mMaxId
    rsOffline!Recorddate = Date
    rsOffline!EndTime = mEndTime22
    rsOffline!StartTime = mTimeStart
    rsOffline!SessionCode = SessionCode
    rsOffline!POSname = POSlServer
    
    rsOffline!CountSalesOfeers = CountSalesOfeers
    rsOffline!CountSales = CountSales
    rsOffline!CountSalesReturn = CountSalesReturn
    rsOffline!CountPurchase = CountPurchase
    rsOffline!CountPurchaseReturn = CountPurchaseReturn
    rsOffline!CountRec = CountRec
    rsOffline.Update
    
    Cn.CommitTrans
    BeginTrans = False
End If

'MsgBox " „ ‰ﬁ· «·»Ì«‰« "

 
    





'Dim mMaxId As Long
's = "Select Max(Id) as MaxID  from TblOffline"
'rsOffline.Open s, Cn, adOpenKeyset, adLockOptimistic
'mMaxId = 1
'If Not rsOffline.EOF Then
'    mMaxId = Val(rsOffline!MaxID & "") + 1
'
'End If
'rsOffline.Close

lblWait.Visible = False

txtEndTime = mEndTime22
txtCountSalesReturn = CountSalesReturn
txtCountSales = CountSales
txtCountSalesOfeers = CountSalesOfeers
Exit Sub



ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'Resume Next
'MsgBox "Done"
'////////////////////////////////////////copy Sales Transactions


End Sub

Private Sub DCboBranch_Click(Area As Integer)
txtBranch = DCboBranch.BoundText
End Sub


'Private Sub Form_Load()
''21 11 2017
'' „  ‰›Ì– «·„»Ì⁄«  ﬂ«„·… „⁄ ﬁÌœÂ«  „⁄ ”‰œ «·’—› „⁄ ﬁÌœ…
''
''
''
' On Error Resume Next
'
'
'txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
'TxtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
'TxtUSERID = GetSetting("ConvertToAccess", "Setting", "USERID", "USERID")
'TxtCHECKTIME = GetSetting("ConvertToAccess", "Setting", "CHECKTIME", "CHECKTIME")
''DcTime.Value = GetSetting("ConvertToAccess", "Setting", "UpdateHours", "00")
'dbRecordDate = Date
''TxtServerDataBaseName = SysSQLServerDataBaseName
''DestinationServer = SysSQLServerName
'POSlServer = "."
'
'
''POSlServer = "PC2\SQL2019"
'
'TxtPOSDB = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")
''POSlServer = SysSQLServerName
''LOCALPOS = SysSQLServerName
''TxtPOSDB = SysSQLServerDataBaseName
'
'txtFromDate.Value = Date
'txtToDate.Value = Date
''BranchDigit = 1
'Dim Msg As String
'
''MsgBox "1"
'If Dir(App.Path & "\pos.txt", vbNormal) = "" Then
'            Msg = "„·›  ”ÃÌ· «·ﬁÊ«⁄œ €Ì— „ÊÃÊœ ...!!!"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'
'           End
'
'        End If
'
'    Open App.Path & "\pos.txt" For Input As #1
'    POSname.Clear
'
'    Do Until EOF(1)
'        Line Input #1, a
'        'subsequent lines
'
'        If a <> "" Then
'            VarSet = Split(a, "*", , vbTextCompare)
'
'            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
'
'                POSname.AddItem (VarSet(0))
'                ServersName.AddItem (VarSet(1))
'                DbName.AddItem (VarSet(2))
'
'            End If
'        End If
'
'    Loop
'   Dim StrSQL As String
'
'TxtServerDataBaseName = VarSet(0)
'   mDBPOSName = VarSet(1)
'
'   'MsgBox "2"
' If ConnectionFirst(True) = False Then
'        Exit Sub
'    End If
'
'        StrSQL = "SELECT branch_id,branch_name FROM TblBranchesData"
'
'
'
'
'
'    GetComboData DCboBranch, StrSQL
'
'    Close #1
'
'POSname_Change
'
''cmdTransfer_Click
''Command11_Click
''
''
''Command2_Click
''Command15_Click
'''Command12_Click
'' 'Command1_Click
''
'' Command10_Click
'End Sub
'
Private Sub Form_Load()

    On Error GoTo errHandler

    Dim Msg As String
    Dim StrSQL As String
    Dim a As String
    Dim VarSet As Variant
    Dim lastValidIndex As Long
    Dim fno As Integer

    '========================
    '  Õ„Ì· «·≈⁄œ«œ«  «·„Õ·Ì… ··‰ﬁÿ…
    '========================
    txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
    TxtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
    TxtUSERID = GetSetting("ConvertToAccess", "Setting", "USERID", "USERID")
    TxtCHECKTIME = GetSetting("ConvertToAccess", "Setting", "CHECKTIME", "CHECKTIME")

    dbRecordDate = Date

    '========================
    ' „Â„ Ãœ«:
    ' «·‰ﬁÿ… «·„Õ·Ì…  ›÷· Local
    ' Õ Ï ·« Ì √À— Command14
    '========================
    POSlServer.Text = "."
    TxtPOSDB.Text = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")

    txtFromDate.Value = Date
    txtToDate.Value = Date

    '========================
    '  ›—Ì€ «·ﬁÊ«∆„
    '========================
    POSname.Clear
    ServersName.Clear
    DbName.Clear

    lastValidIndex = -1

    '========================
    ' «· √ﬂœ „‰ „·› pos.txt
    ' Â–« «·„·› ⁄·Ï «·‰ﬁÿ… ÊÌÕ ÊÌ »Ì«‰«  «·”Ì—›— «·„—ﬂ“Ì
    ' «· ‰”Ìﬁ «·„ Êﬁ⁄:
    ' ServerDb * POSDisplayName * ServerAddress
    ' „À«·:
    ' Test*Test*172.187.248.126,51433*
    '========================
    If Dir(App.Path & "\pos.txt", vbNormal) = "" Then
        Msg = "„·›  ”ÃÌ· «·ﬁÊ«⁄œ pos.txt €Ì— „ÊÃÊœ ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End
    End If

    fno = FreeFile
    Open App.Path & "\pos.txt" For Input As #fno

    Do Until EOF(fno)
        Line Input #fno, a
        a = Trim$(a)

        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            ' ·«“„ ⁄·Ï «·√ﬁ· 3 ⁄‰«’—:
            ' 0 = ServerDb
            ' 1 = «”„ ⁄—÷
            ' 2 = Server Address
            If IsArray(VarSet) Then
                If UBound(VarSet) >= 2 Then
                    If Trim$(VarSet(0) & "") <> "" And Trim$(VarSet(2) & "") <> "" Then

                        ' «”„ «·⁄—÷ ›Ì «·ﬂ„»Ê
                        If Trim$(VarSet(1) & "") <> "" Then
                            POSname.AddItem Trim$(VarSet(1) & "")
                        Else
                            POSname.AddItem Trim$(VarSet(0) & "")
                        End If

                        ' ‰Œ“‰ ﬁ«⁄œ… «·”Ì—›— «·„—ﬂ“Ì
                        DbName.AddItem Trim$(VarSet(0) & "")

                        ' ‰Œ“‰ ⁄‰Ê«‰ «·”Ì—›— «·„—ﬂ“Ì
                        ServersName.AddItem Trim$(VarSet(2) & "")

                        lastValidIndex = POSname.ListCount - 1
                    End If
                End If
            End If
        End If
    Loop

    Close #fno

    If POSname.ListCount = 0 Then
        MsgBox "„·› pos.txt ·« ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ’ÕÌÕ… ··”Ì—›— «·„—ﬂ“Ì", vbCritical, App.Title
        End
    End If

    '========================
    ' «Œ Ì«— √Ê· ”ÿ— «› —«÷Ì«
    '========================
    POSname.ListIndex = 0

    ' ﬁ«⁄œ… «·”Ì—›— «·„—ﬂ“Ì
    TxtServerDataBaseName.Text = DbName.List(POSname.ListIndex)

    ' ⁄‰Ê«‰ «·”Ì—›— «·„—ﬂ“Ì
    SysSQLServerName = ServersName.List(POSname.ListIndex)
    DestinationServer = SysSQLServerName

    ' ··«Õ ›«Ÿ »«· Ê«›ﬁ „⁄ √Ì √ﬂÊ«œ ﬁœÌ„…  ” Œœ„ «·„ €Ì— œÂ
    mDBPOSName = TxtServerDataBaseName.Text

    '========================
    ' › Õ « ’«·  Õ„Ì· ›ﬁÿ
    ' „Â„: IsLoad=True Õ Ï ·« ÌœŒ· ›Ì √Ì  ÂÌ∆«  ≈÷«›Ì…
    ' ﬁœ  ƒÀ— ⁄·Ï «·‰ﬁ·
    '========================
    If ConnectionFirst(True) = False Then
        Exit Sub
    End If

    '========================
    '  Õ„Ì· «·›—Ê⁄ „‰ «·”Ì—›— «·„—ﬂ“Ì
    '========================
    StrSQL = "SELECT branch_id, branch_name FROM TblBranchesData"
    GetComboData DCboBranch, StrSQL

    '========================
    ' «” ﬂ„«·  Õ„Ì· »Ì«‰«  «·‘«‘…
    ' POSname_Change ”Ì⁄„· »‰«¡ ⁄·Ï:
    ' - POSlServer = .
    ' - TxtPOSDB = ﬁ«⁄œ… «·‰ﬁÿ… «·„Õ·Ì…
    ' - SysSQLServerName = «·”Ì—›— «·„—ﬂ“Ì „‰ pos.txt
    ' - TxtServerDataBaseName = ﬁ«⁄œ… «·”Ì—›— «·„—ﬂ“Ì „‰ pos.txt
    '========================
    POSname_Change
cmdTransfer_Click
    Exit Sub

errHandler:
    On Error Resume Next
    If fno <> 0 Then Close #fno

    MsgBox "ÕœÀ Œÿ√ √À‰«¡  Õ„Ì· «·‘«‘…:" & vbCrLf & Err.Description, vbCritical, App.Title
End Sub
 Private Sub GetComboData(My_Combo As DataCombo, _
                         My_SQL As String)
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    Dim Msg As String
    On Error GoTo ErrorHandler

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then
        Exit Sub
    End If

    My_Combo.Tag = My_SQL
    Set rs = New ADODB.Recordset

    
        rs.CursorLocation = adUseClient
   

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly

    'Populate the ADO datacombo by setting its properties
    With My_Combo
        StrTemp = .BoundText
        Set .RowSource = rs
        .BoundColumn = rs(0).Name
        .ListField = rs(1).Name

        If Trim(StrTemp) <> "" Then
            .BoundText = StrTemp
        Else
            .BoundText = ""
            .Text = ""
        End If

    End With

Exit_Sub:
    Set rs = Nothing
    Exit Sub
ErrorHandler:

    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

Private Sub grd_Click()
If grd.Row <> 0 Then
   ' dbRecordDate = grd.TextMatrix(grd.Row, grd.ColIndex("Transaction_Date"))
End If
End Sub

'Private Sub POSname_Change()
'
'  If ConnectionFirst = False Then
'        Exit Sub
'    End If
'    Dim StrSQL As String
'    If POSlServer.Text = "" Then
'        MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
'    Exit Sub
'End If
'
'
'
'   Dim NoOFItem_POS As Double
'   Dim NoOFItem_Server As Double
'
'   Dim Rs3 As New ADODB.Recordset
'   Dim MaxItem_POS As Double
'   Dim MaxItem_Server As Double
'   'step one check item
'
'    ss = "     USE " & ServerDb & vbNewLine
'
'    Cn.Execute ss
'    ss = "USE " & POSDb & vbNewLine
'    POSConnection.Execute ss
'
'    sql = " "
'
'    sql = sql & "     SELECT SUM(CountSales) CountSales ,SUM(CountReturn) CountReturn,SUM(CountSalesOfeers) CountSalesOfeers,Transaction_Date FROM ("
'    sql = sql & "         SELECT COUNT(t.Transaction_ID)     CountTotal,"
'    sql = sql & "                CountSales       = ("
'    sql = sql & "                    Case t.Transaction_Type"
'    sql = sql & "                         WHEN 21 THEN COUNT(t.Transaction_ID)"
'    sql = sql & "                         ELSE 0"
'    sql = sql & "                    End"
'    sql = sql & "                ),"
'    sql = sql & "                CountReturn     = ("
'    sql = sql & "                    Case t.Transaction_Type"
'    sql = sql & "                         WHEN 9 THEN COUNT(t.Transaction_ID)"
'    sql = sql & "                         ELSE 0"
'    sql = sql & "                    End"
'    sql = sql & "                ),"
'
'    sql = sql & "                CountSalesOfeers     = ("
'    sql = sql & "                    Case t.Transaction_Type"
'    sql = sql & "                         WHEN 42 THEN COUNT(t.Transaction_ID)"
'    sql = sql & "                         ELSE 0"
'    sql = sql & "                    End"
'    sql = sql & "                ),"
'
'    sql = sql & "                t.Transaction_Date,"
'    sql = sql & "                Transaction_Type"
'    sql = sql & "         FROM   Transactions             AS t"
'    sql = sql & "         Where IsNull(t.Copied, 0) = 0"
'    sql = sql & "                AND (t.Transaction_Type = 9 OR t.Transaction_Type = 21 )"
'    sql = sql & "         Group By"
'    sql = sql & "                Transaction_Date,"
'    sql = sql & "                Transaction_Type"
'
'    sql = sql & "         ) T"
'    sql = sql & "         Group By"
'    sql = sql & "                Transaction_Date"
'    sql = sql & "         Order By"
'    sql = sql & "                Transaction_Date"
'
'     Text5 = sql
'    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
'    grd.Rows = 1
'    grd.Rows = 2
'    Do While Not Rs3.EOF
'        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountSales")) = Rs3!CountSales & ""
'        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountReturn")) = Rs3!CountReturn & ""
'        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountSalesOfeers")) = Rs3!CountSalesOfeers & ""
'        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("Transaction_Date")) = Rs3!Transaction_Date & ""
'        Rs3.MoveNext
'        grd.Rows = grd.Rows + 1
'    Loop
'    Rs3.Close
'End Sub
'
Private Sub POSname_Change()

    On Error GoTo errHandler

    Dim StrSQL As String
    Dim sql As String
    Dim ss As String

    Dim Rs3 As ADODB.Recordset

    Dim rowIndex As Long
    Dim mCountSales As String
    Dim mCountReturn As String
    Dim mCountOffers As String
    Dim mTranDate As String

    If POSname.ListIndex < 0 Then Exit Sub

    '==================================================
    ' „Â„ Ãœ«:
    ' ·« ‰€Ì¯— «·‰ﬁÿ… «·„Õ·Ì… Â‰«
    ' Õ Ï ·« Ì √À— Command14
    '==================================================
    POSlServer.Text = "."
    TxtPOSDB.Text = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")

    '==================================================
    ' ‰ﬁ—√ »Ì«‰«  «·”Ì—›— «·„—ﬂ“Ì „‰ «·ﬁÊ«∆„ «·„Õ„·… „‰ pos.txt
    '==================================================
    If POSname.ListIndex <= ServersName.ListCount - 1 Then
        SysSQLServerName = Trim$(ServersName.List(POSname.ListIndex) & "")
        DestinationServer = SysSQLServerName
    End If

    If POSname.ListIndex <= DbName.ListCount - 1 Then
        TxtServerDataBaseName.Text = Trim$(DbName.List(POSname.ListIndex) & "")
        mDBPOSName = TxtServerDataBaseName.Text
    End If

    If Trim$(POSlServer.Text) = "" Then
        MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
        Exit Sub
    End If

    If Trim$(TxtPOSDB.Text) = "" Then
        MsgBox "ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ… €Ì— „Õœœ…", vbCritical, "OFFLINE"
        Exit Sub
    End If

    If Trim$(SysSQLServerName) = "" Then
        MsgBox "«”„ √Ê ⁄‰Ê«‰ «·”Ì—›— «·„—ﬂ“Ì €Ì— „Õœœ", vbCritical, "OFFLINE"
        Exit Sub
    End If

    If Trim$(TxtServerDataBaseName.Text) = "" Then
        MsgBox "ﬁ«⁄œ… »Ì«‰«  «·”Ì—›— «·„—ﬂ“Ì €Ì— „Õœœ…", vbCritical, "OFFLINE"
        Exit Sub
    End If

    '==================================================
    ' › Õ «·« ’«·Ì‰:
    ' Cn   = «·”Ì—›— «·„—ﬂ“Ì
    ' POSConnection = «·‰ﬁÿ… «·„Õ·Ì…
    '==================================================
    If ConnectionFirst = False Then
        Exit Sub
    End If

    Set Rs3 = New ADODB.Recordset

    '==================================================
    ' ›ﬁÿ ·· √ﬂœ „‰ ﬁ«⁄œ… «·”Ì—›— Êﬁ«⁄œ… «·‰ﬁÿ…
    '==================================================
    ss = "USE " & TxtServerDataBaseName.Text
    Cn.Execute ss

    ss = "USE " & TxtPOSDB.Text
    POSConnection.Execute ss

    '==================================================
    ' «·≈Õ’«∆Ì… „‰ «·‰ﬁÿ… «·„Õ·Ì… ›ﬁÿ
    ' ⁄‘«‰ ‰⁄—÷ ⁄œœ «·›Ê« Ì— €Ì— «·„‰ﬁÊ·…
    '==================================================
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "       SUM(CountSales) AS CountSales, "
    sql = sql & "       SUM(CountReturn) AS CountReturn, "
    sql = sql & "       SUM(CountSalesOfeers) AS CountSalesOfeers, "
    sql = sql & "       Transaction_Date "
    sql = sql & "FROM ( "
    sql = sql & "    SELECT "
    sql = sql & "           COUNT(t.Transaction_ID) AS CountTotal, "
    sql = sql & "           CASE t.Transaction_Type "
    sql = sql & "                WHEN 21 THEN COUNT(t.Transaction_ID) "
    sql = sql & "                ELSE 0 "
    sql = sql & "           END AS CountSales, "
    sql = sql & "           CASE t.Transaction_Type "
    sql = sql & "                WHEN 9 THEN COUNT(t.Transaction_ID) "
    sql = sql & "                ELSE 0 "
    sql = sql & "           END AS CountReturn, "
    sql = sql & "           CASE t.Transaction_Type "
    sql = sql & "                WHEN 42 THEN COUNT(t.Transaction_ID) "
    sql = sql & "                ELSE 0 "
    sql = sql & "           END AS CountSalesOfeers, "
    sql = sql & "           t.Transaction_Date, "
    sql = sql & "           t.Transaction_Type "
    sql = sql & "    FROM Transactions AS t "
    sql = sql & "    WHERE ISNULL(t.Copied, 0) = 0 "
    sql = sql & "      AND (t.Transaction_Type = 9 OR t.Transaction_Type = 21 OR t.Transaction_Type = 42) "
    sql = sql & "    GROUP BY t.Transaction_Date, t.Transaction_Type "
    sql = sql & ") T "
    sql = sql & "GROUP BY Transaction_Date "
    sql = sql & "ORDER BY Transaction_Date"

    Text5.Text = sql

    Rs3.Open sql, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

    grd.Rows = 1
    grd.Rows = 2

    rowIndex = 1

    Do While Not Rs3.EOF

        If IsNull(Rs3!CountSales) Then
            mCountSales = "0"
        Else
            mCountSales = Rs3!CountSales & ""
        End If

        If IsNull(Rs3!CountReturn) Then
            mCountReturn = "0"
        Else
            mCountReturn = Rs3!CountReturn & ""
        End If

        If IsNull(Rs3!CountSalesOfeers) Then
            mCountOffers = "0"
        Else
            mCountOffers = Rs3!CountSalesOfeers & ""
        End If

        If IsNull(Rs3!Transaction_Date) Then
            mTranDate = ""
        Else
            mTranDate = Rs3!Transaction_Date & ""
        End If

        grd.TextMatrix(rowIndex, grd.ColIndex("CountSales")) = mCountSales
        grd.TextMatrix(rowIndex, grd.ColIndex("CountReturn")) = mCountReturn
        grd.TextMatrix(rowIndex, grd.ColIndex("CountSalesOfeers")) = mCountOffers
        grd.TextMatrix(rowIndex, grd.ColIndex("Transaction_Date")) = mTranDate

        Rs3.MoveNext

        If Not Rs3.EOF Then
            grd.Rows = grd.Rows + 1
            rowIndex = rowIndex + 1
        End If
    Loop

    Rs3.Close
    Set Rs3 = Nothing

    Exit Sub

errHandler:
    On Error Resume Next

    If Not Rs3 Is Nothing Then
        If Rs3.State = adStateOpen Then Rs3.Close
        Set Rs3 = Nothing
    End If

    MsgBox "ÕœÀ Œÿ√ √À‰«¡  Õ„Ì· »Ì«‰«  «·‰ﬁÿ…:" & vbCrLf & Err.Description, vbCritical, App.Title
End Sub
Private Sub POSname_Click()
On Error Resume Next
'    DbName.ListIndex = POSname.ListIndex
'    ServersName.ListIndex = POSname.ListIndex
'
'   POSlServer.text = ServersName.text
'    TxtPOSDB.text = DbName.text
    
    POSname_Change
    
    
    
End Sub
Private Function GetQuery(Optional ByVal mBranchId As Long = 0) As String
    Dim s As String
'    s = "(1 = 1)  "
'    If chkSales.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 21 "
'    End If
'
'    If chkSalesReturn.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 9 "
'    End If
'
'    If chkPurchaseReturn.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 5 "
'    End If
'
'    If chkPurchase.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 22"
'    End If
'
'
'    If chkIn.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 20"
'    End If
'
'    If chkOut.Value = vbChecked Then
'        s = s & " Or Transaction_Type = 19"
'    End If


    Dim tempString As String
    Dim i As Integer
    tempString = "0"
    
    If mBranchId <> 0 Then
        tempString = tempString & "," & 10
        tempString = tempString & "," & 11
        
        
    End If
    If chkSales.Value = vbChecked And mBranchId = 0 Then
        tempString = tempString & "," & 21
'        tempString = tempString & "," & 10
'        tempString = tempString & "," & 11
    End If
    If chkSalesReturn.Value = vbChecked Then
        tempString = tempString & "," & 9
        
    End If
    
    If chkSalesOffers.Value = vbChecked Then
        tempString = tempString & "," & 42
        
        
    End If
    
    
    'GetTransIds = tempString
    
    s = s & "  (Transaction_Type in ( " & tempString & " ) )"
    If mBranchId <> 0 Then
        s = s & " Transactions.StoreID In (Select TblStore.StoreID from TblStore where TblStore.BranchId = " & mBranchId & " )"
    End If
     's = s and dbo.Transactions.Transaction_Date ='" & SQLDate(dbRecordDate.Value, False) & "')"
     
'     s = s & " and (dbo.Transactions.Transaction_Date >='" & SQLDate(txtFromDate.Value, False) & "')"
'     s = s & " and (dbo.Transactions.Transaction_Date <='" & SQLDate(txtToDate.Value, False) & "')"
'
     If chkSalesOffers.Value = vbChecked Then
        s = s & " and  (IsNull(Transactions.DepandToConv,0) = 1  Or Transactions.Transaction_ID In (Select ApprovalData.Transaction_ID from ApprovalData "
        s = s & " where IsNull(ApprovDate,'') <> '' and ScreenName = 'FrmPO1' ))"
        
     End If
    
GetQuery = s
End Function

Private Sub VSFlexGrid1_Click()

End Sub

Private Sub Timer2_Timer()
DoEvents
End Sub

Private Sub txtBranch_Change()
DCboBranch.BoundText = Val(txtBranch)
End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "123" Then
    Frame4.Visible = True
Else
    Frame4.Visible = False
End If
End Sub



Private Function ReserveDestId(ByVal CnX As ADODB.Connection) As String
    On Error GoTo EH
    
    Dim rs As ADODB.Recordset
    Dim s As String
    
    Set rs = New ADODB.Recordset
    
    s = "EXEC dbo.ReserveTransactionId"
    mLastSQL = s
    mLastProc = "ReserveDestId"
    
    rs.Open s, CnX, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF Then
        Err.Raise vbObjectError + 3001, , "Stored procedure dbo.ReserveTransactionId ·„  ı—Ã⁄ √Ì ﬁÌ„…"
    End If
    
    ReserveDestId = Trim$(rs.Fields(0).Value & "")
    
    SafeCloseRS rs
    Exit Function

EH:
    LogAdoErrors CnX, "ReserveDestId", s, "›‘· ›Ì ÕÃ“ Transaction_ID ÃœÌœ"
    ReserveDestId = ""
    SafeCloseRS rs
End Function









Private Sub cmdTransferMove_Click()

    On Error GoTo ErrorHandler

    Dim rsTrans As ADODB.Recordset
    Dim rsDet As ADODB.Recordset
    Dim rsExist As ADODB.Recordset

    Dim FetchSize As Long
    Dim BatchSize As Long
    Dim SessionCode As String

    Dim sql As String
    Dim LastSQL As String

    Dim lastId As Long
    Dim doneAll As Boolean

    Dim transHeader As String
    Dim detHeader As String
    Dim transValuesOnly As String
    Dim detValuesOnly As String

    Dim recCounter As Long
    Dim inTx As Boolean

    Dim mTimeStart As Date
    Dim elapsedSec As Long
    Dim elapsedMin As Long

    Dim newId As String
    Dim v As String
    Dim d As String
    Dim SrcTransId As Long

    mLastProc = "cmdTransferMove_Click"
    mLastSQL = ""
    LastSQL = ""
    inTx = False

    If Trim$(POSlServer.Text) = "" Then
        MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
        Exit Sub
    End If

    If ConnectionFirst = False Then Exit Sub

    lblWait.Visible = True
    lblWait.Caption = "Ì „ «·¬‰ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì…"
    DoEvents

    Set rsTrans = New ADODB.Recordset
    Set rsDet = New ADODB.Recordset
    Set rsExist = New ADODB.Recordset

    rsTrans.CursorLocation = adUseClient
    rsDet.CursorLocation = adUseClient
    rsExist.CursorLocation = adUseClient

    FetchSize = 200
    BatchSize = 5
    SessionCode = Format$(Now, "yyyymmddhhmmss")

    lastId = 0
    doneAll = False
    mTimeStart = Now

    transHeader = ""
    transHeader = transHeader & "INSERT INTO dbo.[Transactions] ("
    transHeader = transHeader & "Transaction_ID, OldTransaction_ID, Transaction_Date, "
    transHeader = transHeader & "Transaction_Serial, Transaction_Type, CusID, StoreID, "
    transHeader = transHeader & "UserID, Emp_ID, BranchId, VAT, VATYou, "
    transHeader = transHeader & "NoteSerial, NoteSerial1, Copied, SessionCode, "
    transHeader = transHeader & "OldNoteserial1, order_no, TaxAddValue, NetValue, "
    transHeader = transHeader & "Transaction_NetValue, Ser"
    transHeader = transHeader & ") VALUES "

    detHeader = ""
    detHeader = detHeader & "INSERT INTO dbo.[Transaction_Details] ("
    detHeader = detHeader & "Transaction_ID, Item_ID, ItemCase, Quantity, Price, "
    detHeader = detHeader & "ItemDiscountType, ItemDiscount, ShowQty, showPrice, "
    detHeader = detHeader & "UnitId, ColorID, ItemSize, ClassId, SessionCode, "
    detHeader = detHeader & "Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, "
    detHeader = detHeader & "Deferred, AmountH, AmountHComm, DetailsPump, "
    detHeader = detHeader & "Account_CodeComm, Account_Code, IsOther"
    detHeader = detHeader & ") VALUES "

    Do While Not doneAll

        lblWait.Caption = "Ì „  Õ„Ì· œ›⁄… ÃœÌœ… „‰ «· ÕÊÌ·«  «·„Œ“‰Ì…... ¬Œ— ID = " & CStr(lastId)
        DoEvents

        sql = ""
        sql = sql & "SELECT TOP (" & FetchSize & ") * "
        sql = sql & "FROM Transactions T2 WITH (READUNCOMMITTED) "
        sql = sql & "WHERE T2.Transaction_Type IN (10,11) "
        sql = sql & "AND T2.Transaction_ID > " & CStr(lastId) & " "
        sql = sql & "ORDER BY T2.Transaction_ID"

        mLastSQL = sql
        LastSQL = sql

        If rsTrans.State <> adStateClosed Then rsTrans.Close
        rsTrans.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rsTrans.EOF Then
            doneAll = True
            Exit Do
        End If

        transValuesOnly = ""
        detValuesOnly = ""
        recCounter = 0

        LastSQL = "SET XACT_ABORT ON;"
        POSConnection.Execute LastSQL, , adExecuteNoRecords

        POSConnection.BeginTrans
        inTx = True

        Do While Not rsTrans.EOF

            SrcTransId = CLng(Val(rsTrans("Transaction_ID").Value & ""))

            If SrcTransId > lastId Then
                lastId = SrcTransId
            End If

            sql = ""
            sql = sql & "SELECT TOP 1 Transaction_ID "
            sql = sql & "FROM Transactions WITH (READUNCOMMITTED) "
            sql = sql & "WHERE OldTransaction_ID = " & CStr(SrcTransId) & " "
            sql = sql & "AND Transaction_Type IN (10,11)"

            If rsExist.State <> adStateClosed Then rsExist.Close
            rsExist.Open sql, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

            If rsExist.EOF Then

                recCounter = recCounter + 1

                newId = ReserveDestId(POSConnection)
                If LenB(newId) = 0 Then
                    Err.Raise vbObjectError + 500, , "·„ Ì „ ÕÃ“ Transaction_ID ÃœÌœ"
                End If

                v = ""
                v = v & "("
                v = v & newId & ","
                v = v & SqlNum(rsTrans("Transaction_ID")) & ","
                v = v & SqlDateTime(rsTrans("Transaction_Date")) & ","
                v = v & SqlStr(rsTrans("Transaction_Serial")) & ","
                v = v & SqlNum(rsTrans("Transaction_Type")) & ","
                v = v & SqlNum(rsTrans("CusID")) & ","
                v = v & SqlNum(rsTrans("StoreID")) & ","
                v = v & SqlNum(rsTrans("UserID")) & ","
                v = v & SqlNum(rsTrans("Emp_ID")) & ","
                v = v & SqlNum(rsTrans("BranchId")) & ","
                v = v & SqlNum(rsTrans("VAT")) & ","
                v = v & SqlNum(rsTrans("VATYou")) & ","
                v = v & SqlStr(rsTrans("NoteSerial")) & ","
                v = v & SqlStr(rsTrans("NoteSerial1")) & ","
                v = v & "1,"
                v = v & SqlStr(SessionCode) & ","
                v = v & SqlStr(rsTrans("OldNoteserial1")) & ","
                v = v & SqlStr(rsTrans("order_no")) & ","
                v = v & SqlNum(rsTrans("TaxAddValue")) & ","
                v = v & SqlNum(rsTrans("NetValue")) & ","
                v = v & SqlNum(rsTrans("Transaction_NetValue")) & ","
                v = v & SqlNum(rsTrans("Ser"))
                v = v & ")"

                If transValuesOnly = "" Then
                    transValuesOnly = v
                Else
                    transValuesOnly = transValuesOnly & "," & vbCrLf & v
                End If

                sql = "SELECT * FROM Transaction_Details WITH (READUNCOMMITTED) WHERE Transaction_ID = " & CStr(SrcTransId)

                If rsDet.State <> adStateClosed Then rsDet.Close
                rsDet.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                Do While Not rsDet.EOF

                    d = ""
                    d = d & "("
                    d = d & newId & ","
                    d = d & SqlNum(rsDet("Item_ID")) & ","
                    d = d & SqlNum(rsDet("ItemCase")) & ","
                    d = d & SqlNum(rsDet("Quantity")) & ","
                    d = d & SqlNum(rsDet("Price")) & ","
                    d = d & SqlNum(rsDet("ItemDiscountType")) & ","
                    d = d & SqlNum(rsDet("ItemDiscount")) & ","
                    d = d & SqlNum(rsDet("ShowQty")) & ","
                    d = d & SqlNum(rsDet("showPrice")) & ","
                    d = d & SqlNum(rsDet("UnitId")) & ","
                    d = d & SqlNum(rsDet("ColorID")) & ","
                    d = d & SqlNum(rsDet("ItemSize")) & ","
                    d = d & SqlNum(rsDet("ClassId")) & ","
                    d = d & SqlStr(SessionCode) & ","
                    d = d & SqlNum(rsDet("Vatyo")) & ","
                    d = d & SqlNum(rsDet("PumpId")) & ","
                    d = d & SqlNum(rsDet("PrevQty")) & ","
                    d = d & SqlStr(rsDet("PrintName")) & ","
                    d = d & SqlNum(rsDet("Cash")) & ","
                    d = d & SqlNum(rsDet("Mada")) & ","
                    d = d & SqlNum(rsDet("Visa")) & ","
                    d = d & SqlNum(rsDet("Deferred")) & ","
                    d = d & SqlNum(rsDet("AmountH")) & ","
                    d = d & SqlNum(rsDet("AmountHComm")) & ","
                    d = d & SqlStr(rsDet("DetailsPump")) & ","
                    d = d & SqlStr(rsDet("Account_CodeComm")) & ","
                    d = d & SqlStr(rsDet("Account_Code")) & ","
                    d = d & SqlNum(rsDet("IsOther"))
                    d = d & ")"

                    If detValuesOnly = "" Then
                        detValuesOnly = d
                    Else
                        detValuesOnly = detValuesOnly & "," & vbCrLf & d
                    End If

                    rsDet.MoveNext
                Loop

                rsDet.Close

                If (recCounter Mod BatchSize) = 0 Then

                    If transValuesOnly <> "" Then
                        LastSQL = transHeader & transValuesOnly
                        mLastSQL = LastSQL
                        POSConnection.Execute LastSQL, , adExecuteNoRecords
                        transValuesOnly = ""
                    End If

                    If detValuesOnly <> "" Then
                        LastSQL = detHeader & detValuesOnly
                        mLastSQL = LastSQL
                        POSConnection.Execute LastSQL, , adExecuteNoRecords
                        detValuesOnly = ""
                    End If

                    POSConnection.CommitTrans
                    inTx = False

                    POSConnection.BeginTrans
                    inTx = True
                End If
            End If

            rsExist.Close
            rsTrans.MoveNext
        Loop

        If transValuesOnly <> "" Then
            LastSQL = transHeader & transValuesOnly
            mLastSQL = LastSQL
            POSConnection.Execute LastSQL, , adExecuteNoRecords
        End If

        If detValuesOnly <> "" Then
            LastSQL = detHeader & detValuesOnly
            mLastSQL = LastSQL
            POSConnection.Execute LastSQL, , adExecuteNoRecords
        End If

        POSConnection.CommitTrans
        inTx = False
    Loop

    elapsedSec = DateDiff("s", mTimeStart, Now)
    elapsedMin = elapsedSec \ 60
    elapsedSec = elapsedSec Mod 60

    lblWait.Caption = " „ «·‰ﬁ·. «·Êﬁ : " & elapsedMin & "œ " & elapsedSec & "À"

    frmPopup.ShowMessage " „ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì… »‰Ã«Õ." & vbCrLf & _
                         "«·Êﬁ : " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì….", vbInformation

CleanExit:
    lblWait.Visible = False
    SafeCloseRS rsExist
    SafeCloseRS rsDet
    SafeCloseRS rsTrans
    Exit Sub

ErrorHandler:
    On Error Resume Next

    If inTx Then
        POSConnection.RollbackTrans
    End If

    LogAdoErrors POSConnection, "cmdTransferMove_Click", LastSQL, "Œÿ√ √À‰«¡ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì…"

    frmPopup.ShowMessage BuildErrMsg(POSConnection, "cmdTransferMove_Click", LastSQL, "ÕœÀ Œÿ√ √À‰«¡ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì…"), vbCritical

    lblWait.Visible = False
    SafeCloseRS rsExist
    SafeCloseRS rsDet
    SafeCloseRS rsTrans
    Err.Clear

End Sub





'Private Sub cmdTransferMove_Click()
'  On Error GoTo ErrorHandler
'
'  If POSlServer.Text = "" Then
'      MsgBox "«Œ — «·‰ﬁÿ… «·„ ’·… √Ê·«", vbCritical, "Œÿ√"
'      Exit Sub
'  End If
'  If ConnectionFirst = False Then Exit Sub
'
'  lblWait.Visible = True
'  lblWait.Caption = "Ì „ «·«‰ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì…"
'  DoEvents
'
'  '==== « ’«· «·‰ﬁÿ… ====
'  Dim POSConnection As New ADODB.Connection
'  POSConnection.CursorLocation = adUseServer
'  POSConnection.ConnectionString = _
'      "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & _
'      ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'      ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
'  POSConnection.Open
'  POSConnection.CommandTimeout = 0  ' ·« „Â·…
'
'  '==== Prefixes ====
'  '  Linked server ··”Ì—›— «·„—ﬂ“Ì
'  Dim mServerD As String, mPosD As String
'  mPosD = "dbo."                                                 ' ·√‰ «·‹ Initial Catalog = Nazim
'  mServerD = "[" & "RemoteServer10" & "]." & mDBPOSName & ".dbo." ' «”„ «··Ì‰ﬂœ + DB «·”Ì—›—
'
'  '==== ≈⁄œ«œ«  ====
'  Dim FetchSize As Long: FetchSize = 200   ' ÕÃ„ «·’›Õ…
'  Dim BatchSize As Long: BatchSize = 5     ' ‰ﬂ„¯  ﬂ· 5  ÕÊÌ·« 
'  Dim SessionCode As String: SessionCode = Format(Now, "yyyymmddhhmmss")
'
'  Dim sql As String, LastSQL As String
'  Dim lastId As Long: lastId = 0
'  Dim doneAll As Boolean: doneAll = False
'
'  Dim rsTrans As New ADODB.Recordset
'  rsTrans.CursorLocation = adUseServer
'
'  Dim transHeader As String, detHeader As String
'  Dim transValuesOnly As String, detSelects As String
'  Dim recCounter As Long, inTx As Boolean
'
'  transHeader = "INSERT INTO " & mPosD & "[Transactions] (" & _
'                "Transaction_ID, OldTransaction_ID, Transaction_Date, " & _
'                "Transaction_Serial, Transaction_Type, CusID, StoreID, UserID, Emp_ID, BranchId, VAT, VATYou, " & _
'                "NoteSerial, NoteSerial1, Copied, SessionCode, OldNoteserial1, " & _
'                "order_no, TaxAddValue, NetValue, Transaction_NetValue, Ser) VALUES "
'
'  detHeader = "INSERT INTO " & mPosD & "[Transaction_Details] (" & _
'              "Transaction_ID, Item_ID, ItemCase, Quantity, Price, " & _
'              "ItemDiscountType, ItemDiscount, ShowQty, showPrice, UnitId, ColorID, ItemSize, ClassId, " & _
'              "SessionCode, Vatyo, PumpId, PrevQty, PrintName, Cash, Mada, Visa, Deferred, " & _
'              "AmountH, AmountHComm, DetailsPump, Account_CodeComm, Account_Code, IsOther) "
'
'  Dim mTimeStart As Date: mTimeStart = Now
'
'  Do While Not doneAll
'      ' ’›Õ… ÃœÌœ… „‰ «·”Ì—›— «·„—ﬂ“Ì ⁄»— «·‹Linked Server
'      sql = "SELECT TOP (" & FetchSize & ") * " & _
'            "FROM " & mServerD & "Transactions T2 WITH (READUNCOMMITTED) " & _
'            "WHERE T2.Transaction_Type IN (10,11) " & _
'            "  AND T2.Transaction_ID > " & CStr(lastId) & " " & _
'            "  AND NOT EXISTS ( " & _
'            "        SELECT 1 FROM " & mPosD & "Transactions T1 WITH (READUNCOMMITTED) " & _
'            "        WHERE T1.OldTransaction_ID = T2.Transaction_ID " & _
'            "          AND T1.Transaction_Type IN (10,11) ) " & _
'            "ORDER BY T2.Transaction_ID"
'
'      If rsTrans.State <> adStateClosed Then rsTrans.Close
'      rsTrans.CursorLocation = adUseClient
'      'rsTrans.Open sql, POSConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'      rsTrans.Open sql, POSConnection, adOpenStatic, adLockReadOnly, adCmdText
'      Set rsTrans.ActiveConnection = Nothing
'      If rsTrans.EOF Then
'          doneAll = True
'          Exit Do
'      End If
'
'      ' »œ¡ œ›⁄… ÃœÌœ…
'      transValuesOnly = "": detSelects = "": recCounter = 0
'      POSConnection.Execute "SET XACT_ABORT ON;"
'      POSConnection.BeginTrans: inTx = True
'
'      Do While Not rsTrans.EOF
'          recCounter = recCounter + 1
'
'          ' ¬Œ— ID ›Ì «·’›Õ… (·‹ paging)
'          If CLng(rsTrans("Transaction_ID")) > lastId Then lastId = CLng(rsTrans("Transaction_ID"))
'
'          ' Reserve ID ⁄·Ï ‰›” «·‹POSConnection
'          Dim newId As String
'          newId = ReserveDestId(POSConnection)
'          If LenB(newId) = 0 Then Err.Raise vbObjectError + 500, , "·„ Ì „ ÕÃ“ Transaction_ID ÃœÌœ"
'
'          ' ’› «·—ƒÊ”
'          Dim v As String
'          v = "(" & _
'              newId & "," & _
'              SqlNum(rsTrans("Transaction_ID")) & "," & _
'              SqlDateTime(rsTrans("Transaction_Date")) & "," & _
'              SqlStr(rsTrans("Transaction_Serial")) & "," & _
'              SqlNum(rsTrans("Transaction_Type")) & "," & _
'              SqlNum(rsTrans("CusID")) & "," & _
'              SqlNum(rsTrans("StoreID")) & "," & _
'              SqlNum(rsTrans("UserID")) & "," & _
'              SqlNum(rsTrans("Emp_ID")) & "," & _
'              SqlNum(rsTrans("BranchId")) & "," & _
'              SqlNum(rsTrans("VAT")) & "," & _
'              SqlNum(rsTrans("VATYou")) & "," & _
'              SqlStr(rsTrans("NoteSerial")) & "," & _
'              SqlStr(rsTrans("NoteSerial1")) & "," & _
'              "1," & SqlStr(SessionCode) & "," & _
'              SqlStr(rsTrans("OldNoteserial1")) & "," & _
'              SqlStr(rsTrans("order_no")) & "," & _
'              SqlNum(rsTrans("TaxAddValue")) & "," & _
'              SqlNum(rsTrans("NetValue")) & "," & _
'              SqlNum(rsTrans("Transaction_NetValue")) & "," & _
'              SqlNum(rsTrans("Ser")) & ")"
'
'          If transValuesOnly = "" Then transValuesOnly = v Else transValuesOnly = transValuesOnly & "," & vbCrLf & v
'
'          '  ›«’Ì· «· ÕÊÌ·«  (INSERT .. SELECT „‰ «·”Ì—›—)
'          Dim s As String
'          s = "SELECT " & newId & " AS Transaction_ID, " & _
'              "TD.Item_ID, TD.ItemCase, TD.Quantity, TD.Price, " & _
'              "TD.ItemDiscountType, TD.ItemDiscount, TD.ShowQty, TD.showPrice, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId, " & _
'              SqlStr(SessionCode) & " AS SessionCode, TD.Vatyo, TD.PumpId, TD.PrevQty, TD.PrintName, TD.Cash, TD.Mada, TD.Visa, TD.Deferred, " & _
'              "TD.AmountH, TD.AmountHComm, TD.DetailsPump, TD.Account_CodeComm, TD.Account_Code, TD.IsOther " & _
'              "FROM " & mServerD & "Transaction_Details TD WITH (READUNCOMMITTED) " & _
'              "WHERE TD.Transaction_ID = " & SqlNum(rsTrans("Transaction_ID"))
'
'          If detSelects = "" Then detSelects = s Else detSelects = detSelects & vbCrLf & "UNION ALL" & vbCrLf & s
'
'          ' ‰‰›¯– ﬂ· batchSize  ÕÊÌ·«  Ê‰ﬂÊ„ˆ¯ 
'          If (recCounter Mod BatchSize) = 0 Then
'              If transValuesOnly <> "" Then
'                  LastSQL = transHeader & transValuesOnly
'                  POSConnection.Execute LastSQL, , adExecuteNoRecords
'                  transValuesOnly = ""
'              End If
'              If detSelects <> "" Then
'                  LastSQL = detHeader & detSelects
'                  POSConnection.Execute LastSQL, , adExecuteNoRecords
'                  detSelects = ""
'              End If
'
'              POSConnection.CommitTrans: inTx = False
'              POSConnection.BeginTrans:  inTx = True
'          End If
'
'          rsTrans.MoveNext
'      Loop
'
'      ' »«ﬁÌ «·œ›⁄…
'      If transValuesOnly <> "" Then
'          LastSQL = transHeader & transValuesOnly
'          POSConnection.Execute LastSQL, , adExecuteNoRecords
'      End If
'      If detSelects <> "" Then
'          LastSQL = detHeader & detSelects
'          POSConnection.Execute LastSQL, , adExecuteNoRecords
'      End If
'
'      POSConnection.CommitTrans: inTx = False
'  Loop
'
'  '  ﬁ—Ì— Êﬁ 
'  Dim elapsedSec As Long, elapsedMin As Long
'  elapsedSec = DateDiff("s", mTimeStart, Now)
'  elapsedMin = elapsedSec \ 60: elapsedSec = elapsedSec Mod 60
'
'  frmPopup.ShowMessage " „ ‰ﬁ· «· ÕÊÌ·«  «·„Œ“‰Ì… »‰Ã«Õ." & vbCrLf & _
'         "«·Êﬁ : " & elapsedMin & " œﬁÌﬁ… " & elapsedSec & " À«‰Ì….", vbInformation
'  lblWait.Caption = " „ «·‰ﬁ·. «·Êﬁ : " & elapsedMin & "œ " & elapsedSec & "À"
'  Exit Sub
'
'ErrorHandler:
'  On Error Resume Next
'  If inTx Then POSConnection.RollbackTrans
'  frmPopup.ShowMessage "Œÿ√ √À‰«¡ «·‰ﬁ·: " & Err.Description, vbCritical
'End Sub
''==================== œÊ«· „”«⁄œ… »”Ìÿ… ====================
'
'
'

'=== Helpers „Œ ’—… ===
Private Function Esc(ByVal s As String, Optional ByVal maxLen As Long = 2000) As String
    If LenB(s) = 0 Then
        Esc = ""
    Else
        Esc = Left$(Replace(s, "'", "''"), maxLen)
    End If
End Function

Private Sub SyncGroups_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT GroupID, GroupName FROM Groups", Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 GroupID FROM Groups WHERE GroupID = " & SqlNum(rsSrc("GroupID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            s = "(" & SqlNum(rsSrc("GroupID")) & "," & SqlStr(rsSrc("GroupName")) & ")"
            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO Groups (GroupID, GroupName) VALUES ", s, 200
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub
Private Sub SyncTblUnites_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT UnitID, UnitName, UnitNamee FROM TblUnites", Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 UnitID FROM TblUnites WHERE UnitID = " & SqlNum(rsSrc("UnitID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & SqlNum(rsSrc("UnitID")) & "," & SqlStr(rsSrc("UnitName")) & "," & SqlStr(rsSrc("UnitNamee")) & ")"
            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblUnites (UnitID, UnitName, UnitNamee) VALUES ", v, 200
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub
Private Sub SyncTblPaymentType_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT PaymentID, PaymentName, PaymentNamee, Accountcom, commision, branch_no, " & _
               "TaxTobacco, AccTaxTobacco, IsNewCode, IsHiddenVat, IsDefault FROM TblPaymentType", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 PaymentID FROM TblPaymentType WHERE PaymentID = " & SqlNum(rsSrc("PaymentID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("PaymentID")) & "," & _
                SqlStr(rsSrc("PaymentName")) & "," & _
                SqlStr(rsSrc("PaymentNamee")) & "," & _
                SqlStr(rsSrc("Accountcom")) & "," & _
                SqlNum(rsSrc("commision")) & "," & _
                SqlNum(rsSrc("branch_no")) & "," & _
                SqlNum(rsSrc("TaxTobacco")) & "," & _
                SqlStr(rsSrc("AccTaxTobacco")) & "," & _
                SqlNum(rsSrc("IsNewCode")) & "," & _
                SqlNum(rsSrc("IsHiddenVat")) & "," & _
                SqlNum(rsSrc("IsDefault")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblPaymentType (" & _
                "PaymentID, PaymentName, PaymentNamee, Accountcom, commision, branch_no, TaxTobacco, AccTaxTobacco, IsNewCode, IsHiddenVat, IsDefault" & _
                ") VALUES ", v, 200
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblPaymentUser_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT PaynetID, UserID FROM TblPaymentUser", Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 PaynetID FROM TblPaymentUser WHERE ISNULL(PaynetID,0)=" & _
            SqlNum(rsSrc("PaynetID")) & " AND ISNULL(UserID,0)=" & SqlNum(rsSrc("UserID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & SqlNum(rsSrc("PaynetID")) & "," & SqlNum(rsSrc("UserID")) & ")"
            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblPaymentUser (PaynetID, UserID) VALUES ", v, 300
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncBanksData_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT BankID, BankName, BankNamee, Account_Code, Account_Code1, Account_Code2, BranchId, ParetnAccount, parent_account FROM BanksData", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 BankID FROM BanksData WHERE BankID = " & SqlNum(rsSrc("BankID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("BankID")) & "," & _
                SqlStr(rsSrc("BankName")) & "," & _
                SqlStr(rsSrc("BankNamee")) & "," & _
                SqlStr(rsSrc("Account_Code")) & "," & _
                SqlStr(rsSrc("Account_Code1")) & "," & _
                SqlStr(rsSrc("Account_Code2")) & "," & _
                SqlNum(rsSrc("BranchId")) & "," & _
                SqlStr(rsSrc("ParetnAccount")) & "," & _
                SqlStr(rsSrc("parent_account")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO BanksData (BankID, BankName, BankNamee, Account_Code, Account_Code1, Account_Code2, BranchId, ParetnAccount, parent_account) VALUES ", _
                v, 200
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblUsers_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT UserID, UserName, PassWord, BranchId, BoxID, BankID, Empid, FixedCustomer FROM TblUsers", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 UserID FROM TblUsers WHERE UserID = " & SqlNum(rsSrc("UserID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("UserID")) & "," & _
                SqlStr(rsSrc("UserName")) & "," & _
                SqlStr(rsSrc("PassWord")) & "," & _
                SqlNum(rsSrc("BranchId")) & "," & _
                SqlNum(rsSrc("BoxID")) & "," & _
                SqlNum(rsSrc("BankID")) & "," & _
                SqlNum(rsSrc("Empid")) & "," & _
                SqlNum(rsSrc("FixedCustomer")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblUsers (UserID, UserName, PassWord, BranchId, BoxID, BankID, Empid, FixedCustomer) VALUES ", _
                v, 200
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblEmpJobsTypes_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT JobTypeID, JobTypeName, JobTypeNamee FROM TblEmpJobsTypes", Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 JobTypeID FROM TblEmpJobsTypes WHERE JobTypeID = " & SqlNum(rsSrc("JobTypeID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & SqlNum(rsSrc("JobTypeID")) & "," & SqlStr(rsSrc("JobTypeName")) & "," & SqlStr(rsSrc("JobTypeNamee")) & ")"
            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblEmpJobsTypes (JobTypeID, JobTypeName, JobTypeNamee) VALUES ", v, 200
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblEmployee_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT Emp_ID, Emp_Code, Emp_Name, Nationality, dean, JobTypeID, placeWORK, DepartmentID, " & _
               "Emp_Salary, Emp_Salary_others, NumEkama, DateEndekamah, KafelID, NumPasp, jopstatusid, " & _
               "Emp_mobile, BranchId, Emp_Namee FROM TblEmployee", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 Emp_ID FROM TblEmployee WHERE Emp_ID = " & SqlNum(rsSrc("Emp_ID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("Emp_ID")) & "," & _
                SqlStr(rsSrc("Emp_Code")) & "," & _
                SqlStr(rsSrc("Emp_Name")) & "," & _
                SqlStr(rsSrc("Nationality")) & "," & _
                SqlStr(rsSrc("dean")) & "," & _
                SqlNum(rsSrc("JobTypeID")) & "," & _
                SqlStr(rsSrc("placeWORK")) & "," & _
                SqlNum(rsSrc("DepartmentID")) & "," & _
                SqlNum(rsSrc("Emp_Salary")) & "," & _
                SqlNum(rsSrc("Emp_Salary_others")) & "," & _
                SqlStr(rsSrc("NumEkama")) & "," & _
                SqlDateTime(rsSrc("DateEndekamah")) & "," & _
                SqlNum(rsSrc("KafelID")) & "," & _
                SqlStr(rsSrc("NumPasp")) & "," & _
                SqlNum(rsSrc("jopstatusid")) & "," & _
                SqlStr(rsSrc("Emp_mobile")) & "," & _
                SqlNum(rsSrc("BranchId")) & "," & _
                SqlStr(rsSrc("Emp_Namee")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblEmployee (" & _
                "Emp_ID, Emp_Code, Emp_Name, Nationality, dean, JobTypeID, placeWORK, DepartmentID, " & _
                "Emp_Salary, Emp_Salary_others, NumEkama, DateEndekamah, KafelID, NumPasp, jopstatusid, Emp_mobile, BranchId, Emp_Namee" & _
                ") VALUES ", v, 100
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblCustemers_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT CusID, CusName, CusNamee, ResponsibleContact, Cus_mobile, Type, OpenBalance, Account_Code, " & _
               "CityID, EmpId, Address, parent_account, prifix, Fullcode, BranchId, VATNO, CustGID FROM TblCustemers", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 CusID FROM TblCustemers WHERE CusID = " & SqlNum(rsSrc("CusID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("CusID")) & "," & _
                SqlStr(rsSrc("CusName")) & "," & _
                SqlStr(rsSrc("CusNamee")) & "," & _
                SqlStr(rsSrc("ResponsibleContact")) & "," & _
                SqlStr(rsSrc("Cus_mobile")) & "," & _
                SqlNum(rsSrc("Type")) & "," & _
                SqlNum(rsSrc("OpenBalance")) & "," & _
                SqlStr(rsSrc("Account_Code")) & "," & _
                SqlNum(rsSrc("CityID")) & "," & _
                SqlNum(rsSrc("EmpId")) & "," & _
                SqlStr(rsSrc("Address")) & "," & _
                SqlStr(rsSrc("parent_account")) & "," & _
                SqlStr(rsSrc("prifix")) & "," & _
                SqlStr(rsSrc("Fullcode")) & "," & _
                SqlNum(rsSrc("BranchId")) & "," & _
                SqlStr(rsSrc("VATNO")) & "," & _
                SqlNum(rsSrc("CustGID")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblCustemers (" & _
                "CusID, CusName, CusNamee, ResponsibleContact, Cus_mobile, Type, OpenBalance, Account_Code, CityID, EmpId, Address, parent_account, prifix, Fullcode, BranchId, VATNO, CustGID" & _
                ") VALUES ", v, 100
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblStore_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT StoreID, StoreName, Account_Code, Account_Code1, Account_Code2, Emp_ID, Account_Code3, linked, BranchId, " & _
               "Code, StoreNamee, ParetnAccount, SalesPersonId, PurchasePersonid, Account_Code0, Account_Code11, " & _
               "Account_Code22, Account_Code33, BoxID FROM TblStore", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 StoreID FROM TblStore WHERE StoreID = " & SqlNum(rsSrc("StoreID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("StoreID")) & "," & _
                SqlStr(rsSrc("StoreName")) & "," & _
                SqlStr(rsSrc("Account_Code")) & "," & _
                SqlStr(rsSrc("Account_Code1")) & "," & _
                SqlStr(rsSrc("Account_Code2")) & "," & _
                SqlNum(rsSrc("Emp_ID")) & "," & _
                SqlStr(rsSrc("Account_Code3")) & "," & _
                SqlNum(rsSrc("linked")) & "," & _
                SqlNum(rsSrc("BranchId")) & "," & _
                SqlStr(rsSrc("Code")) & "," & _
                SqlStr(rsSrc("StoreNamee")) & "," & _
                SqlStr(rsSrc("ParetnAccount")) & "," & _
                SqlNum(rsSrc("SalesPersonId")) & "," & _
                SqlNum(rsSrc("PurchasePersonid")) & "," & _
                SqlStr(rsSrc("Account_Code0")) & "," & _
                SqlStr(rsSrc("Account_Code11")) & "," & _
                SqlStr(rsSrc("Account_Code22")) & "," & _
                SqlStr(rsSrc("Account_Code33")) & "," & _
                SqlNum(rsSrc("BoxID")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblStore (" & _
                "StoreID, StoreName, Account_Code, Account_Code1, Account_Code2, Emp_ID, Account_Code3, linked, BranchId, Code, StoreNamee, ParetnAccount, SalesPersonId, PurchasePersonid, Account_Code0, Account_Code11, Account_Code22, Account_Code33, BoxID" & _
                ") VALUES ", v, 100
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub SyncTblItemsUnits_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    rsSrc.Open "SELECT JunckID, ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, " & _
               "FactorByDefaultUnit, MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2 FROM TblItemsUnits", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 ItemID FROM TblItemsUnits WHERE ItemID = " & SqlNum(rsSrc("ItemID")) & _
            " AND UnitID = " & SqlNum(rsSrc("UnitID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then
            v = "(" & _
                SqlNum(rsSrc("JunckID")) & "," & _
                SqlNum(rsSrc("ItemID")) & "," & _
                SqlNum(rsSrc("UnitID")) & "," & _
                SqlNum(rsSrc("UnitFactor")) & "," & _
                SqlNum(rsSrc("SecOrder")) & "," & _
                SqlNum(rsSrc("DefaultUnit")) & "," & _
                SqlNum(rsSrc("UnitSalesPrice")) & "," & _
                SqlNum(rsSrc("UnitPurPrice")) & "," & _
                SqlNum(rsSrc("FactorByDefaultUnit")) & "," & _
                SqlNum(rsSrc("MinSelingPrice")) & "," & _
                SqlNum(rsSrc("ForUnit")) & "," & _
                SqlNum(rsSrc("MethodCalc")) & "," & _
                SqlStr(rsSrc("SessionCode")) & "," & _
                SqlStr(rsSrc("barCodeNo2")) & ")"

            AppendBatchInsert BatchSQL, BatchCount, _
                "INSERT INTO TblItemsUnits (" & _
                "JunckID, ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2" & _
                ") VALUES ", v, 150
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc
End Sub

Private Sub UpdateTblItemsUnits_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim s As String

    Set rsSrc = New ADODB.Recordset

    rsSrc.Open "SELECT ItemID, UnitID, UnitSalesPrice, barCodeNo2, MaxSelingPrice, UnitWholeSalePrice, MinSelingPrice, UnitPurPrice FROM TblItemsUnits", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = ""
        s = s & "UPDATE TblItemsUnits SET "
        s = s & "UnitSalesPrice = " & SqlNum(rsSrc("UnitSalesPrice")) & ", "
        s = s & "barCodeNo2 = " & SqlStr(rsSrc("barCodeNo2")) & ", "
        s = s & "MaxSelingPrice = " & SqlNum(rsSrc("MaxSelingPrice")) & ", "
        s = s & "UnitWholeSalePrice = " & SqlNum(rsSrc("UnitWholeSalePrice")) & ", "
        s = s & "MinSelingPrice = " & SqlNum(rsSrc("MinSelingPrice")) & ", "
        s = s & "UnitPurPrice = " & SqlNum(rsSrc("UnitPurPrice")) & " "
        s = s & "WHERE ItemID = " & SqlNum(rsSrc("ItemID")) & " "
        s = s & "AND UnitID = " & SqlNum(rsSrc("UnitID")) & ";"

        AppendBatchUpdate BatchSQL, BatchCount, s, 200
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL
    SafeCloseRS rsSrc
End Sub

Private Sub UpdateTblItems_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim s As String

    Set rsSrc = New ADODB.Recordset

    rsSrc.Open "SELECT ItemID, ItemName, barCodeNO, Code, Fullcode, IsArchive FROM TblItems", _
               Cn, adOpenStatic, adLockReadOnly, adCmdText

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = ""
        s = s & "UPDATE TblItems SET "
        s = s & "ItemName = " & SqlStr(rsSrc("ItemName")) & ", "
        s = s & "barCodeNO = " & SqlStr(rsSrc("barCodeNO")) & ", "
        s = s & "Code = " & SqlStr(rsSrc("Code")) & ", "
        s = s & "Fullcode = " & SqlStr(rsSrc("Fullcode")) & ", "
        s = s & "IsArchive = " & SqlNum(rsSrc("IsArchive")) & " "
        s = s & "WHERE ItemID = " & SqlNum(rsSrc("ItemID")) & ";"

        AppendBatchUpdate BatchSQL, BatchCount, s, 200
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL
    SafeCloseRS rsSrc
End Sub
Private Sub AppendBatchInsert(ByRef BatchSQL As String, ByRef BatchCount As Long, ByVal insertPrefix As String, ByVal ValueRow As String, ByVal FlushEvery As Long)

    If BatchSQL = "" Then
        BatchSQL = insertPrefix & ValueRow
    Else
        BatchSQL = BatchSQL & "," & vbCrLf & ValueRow
    End If

    BatchCount = BatchCount + 1

    If BatchCount >= FlushEvery Then
        FlushBatchInsert POSConnection, BatchSQL
        BatchCount = 0
    End If
End Sub

Private Sub AppendBatchUpdate(ByRef BatchSQL As String, ByRef BatchCount As Long, ByVal SqlLine As String, ByVal FlushEvery As Long)

    If BatchSQL = "" Then
        BatchSQL = SqlLine
    Else
        BatchSQL = BatchSQL & vbCrLf & SqlLine
    End If

    BatchCount = BatchCount + 1

    If BatchCount >= FlushEvery Then
        FlushBatchInsert POSConnection, BatchSQL
        BatchCount = 0
    End If
End Sub

Private Sub FlushBatchInsert(ByVal Cnn As ADODB.Connection, ByRef BatchSQL As String)

    If Trim$(BatchSQL) <> "" Then
        mLastSQL = BatchSQL
        Cnn.Execute BatchSQL, , adExecuteNoRecords
        BatchSQL = ""
    End If
End Sub
Private Sub SyncTblItems_ServerToPOS(ByRef BatchSQL As String, ByRef BatchCount As Long)

    Dim rsSrc As ADODB.Recordset
    Dim rsChk As ADODB.Recordset
    Dim s As String
    Dim v As String
    Dim sqlSrc As String
    Dim insertPrefix As String

    Set rsSrc = New ADODB.Recordset
    Set rsChk = New ADODB.Recordset

    sqlSrc = ""
    sqlSrc = sqlSrc & "SELECT "
    sqlSrc = sqlSrc & "ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, "
    sqlSrc = sqlSrc & "PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, "
    sqlSrc = sqlSrc & "HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, "
    sqlSrc = sqlSrc & "ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, "
    sqlSrc = sqlSrc & "ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode, "
    sqlSrc = sqlSrc & "prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, "
    sqlSrc = sqlSrc & "itemSerials, barCodeNO, SizeID11 "
    sqlSrc = sqlSrc & "FROM TblItems"

    rsSrc.Open sqlSrc, Cn, adOpenStatic, adLockReadOnly, adCmdText

    insertPrefix = ""
    insertPrefix = insertPrefix & "INSERT INTO TblItems ("
    insertPrefix = insertPrefix & "ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, "
    insertPrefix = insertPrefix & "PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, "
    insertPrefix = insertPrefix & "HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, "
    insertPrefix = insertPrefix & "ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, "
    insertPrefix = insertPrefix & "ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode, "
    insertPrefix = insertPrefix & "prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, "
    insertPrefix = insertPrefix & "itemSerials, barCodeNO, SizeID11"
    insertPrefix = insertPrefix & ") VALUES "

    BatchSQL = ""
    BatchCount = 0

    Do While Not rsSrc.EOF

        s = "SELECT TOP 1 ItemID FROM TblItems WHERE ItemID = " & SqlNum(rsSrc("ItemID"))
        rsChk.Open s, POSConnection, adOpenStatic, adLockReadOnly, adCmdText

        If rsChk.EOF Then

            v = ""
            v = v & "("
            v = v & SqlNum(rsSrc("ItemID")) & ","
            v = v & SqlStr(rsSrc("ItemCode")) & ","
            v = v & SqlStr(rsSrc("ItemName")) & ","
            v = v & SqlNum(rsSrc("GroupID")) & ","
            v = v & SqlNum(rsSrc("HaveSerial")) & ","
            v = v & SqlDateTime(rsSrc("LastUpdate")) & ","
            v = v & SqlNum(rsSrc("PurchasePrice")) & ","
            v = v & SqlNum(rsSrc("SallingPrice")) & ","
            v = v & SqlNum(rsSrc("RequestLimit")) & ","
            v = v & SqlNum(rsSrc("CustomerPrice")) & ","
            v = v & SqlNum(rsSrc("HaveGuarantee")) & ","
            v = v & SqlNum(rsSrc("GuaranteeValue")) & ","
            v = v & SqlNum(rsSrc("GuaranteeType")) & ","
            v = v & SqlNum(rsSrc("IsArchive")) & ","
            v = v & SqlNum(rsSrc("ItemType")) & ","
            v = v & SqlNum(rsSrc("AssbliedItem")) & ","
            v = v & SqlNum(rsSrc("RelatedItem")) & ","
            v = v & SqlStr(rsSrc("ItemComment")) & ","
            v = v & SqlNum(rsSrc("ItemCase")) & ","
            v = v & SqlNum(rsSrc("ItemMaking")) & ","
            v = v & SqlNum(rsSrc("ItemMakingNew")) & ","
            v = v & SqlStr(rsSrc("code")) & ","
            v = v & SqlNum(rsSrc("Branch_NO")) & ","
            v = v & SqlStr(rsSrc("Fullcode")) & ","
            v = v & SqlStr(rsSrc("prifix")) & ","
            v = v & SqlStr(rsSrc("PartNo")) & ","
            v = v & SqlNum(rsSrc("CostPrice")) & ","
            v = v & SqlStr(rsSrc("ItemNamee")) & ","
            v = v & SqlNum(rsSrc("DefaultSupplier")) & ","
            v = v & SqlStr(rsSrc("itemSerials")) & ","
            v = v & SqlStr(rsSrc("barCodeNO")) & ","
            v = v & SqlNum(rsSrc("SizeID11"))
            v = v & ")"

            AppendBatchInsert BatchSQL, BatchCount, insertPrefix, v, 100
        End If

        rsChk.Close
        rsSrc.MoveNext
    Loop

    FlushBatchInsert POSConnection, BatchSQL

    SafeCloseRS rsChk
    SafeCloseRS rsSrc

End Sub
Private Function FmtNum(ByVal v As Variant) As String
    If IsNull(v) Or v = "" Then
        FmtNum = "NULL"
    Else
        FmtNum = Replace(CStr(v), ",", ".")
    End If
End Function

Private Function FmtBit(ByVal b As Boolean) As String
    FmtBit = IIf(b, "1", "0")
End Function

'=== «·œ«·… «·„÷€Êÿ… ===
Private Sub SaveSyncLog( _
    ByVal cnPOS As ADODB.Connection, _
    ByVal cnSrv As ADODB.Connection, _
    ByVal SessionCode As String, _
    ByVal direction As String, _
    ByVal TransferKind As String, _
    ByVal StartTime As Date, _
    ByVal EndTime As Date, _
    ByVal DurationSec As Long, _
    ByVal SourceServer As String, ByVal SourceDb As String, _
    ByVal DestServer As String, ByVal DestDb As String, _
    ByVal BranchID As Long, _
    ByVal FilterUsed As String, _
    ByVal BatchSize As Long, ByVal FetchSize As Long, _
    ByVal SrcHeads As Long, ByVal DstHeads As Long, _
    ByVal SrcDet As Long, ByVal DstDet As Long, _
    ByVal SrcVAT As Long, ByVal DstVAT As Long, _
    ByVal SrcPay As Long, ByVal DstPay As Long, _
    ByVal SrcPay2 As Long, ByVal DstPay2 As Long, _
    ByVal SrcAmount As Currency, ByVal DstAmount As Currency, _
    ByVal SrcVATSum As Currency, ByVal DstVATSum As Currency, _
    ByVal SrcTPay As Currency, ByVal DstTPay As Currency, _
    ByVal SrcSPay As Currency, ByVal DstSPay As Currency, _
    ByVal IsOk As Boolean, _
    ByVal ErrorMsg As String)

    On Error Resume Next

    Dim sql As String
    sql = "INSERT INTO dbo.SyncTransferLog (" _
        & "SessionCode,Direction,TransferKind,StartTime,EndTime,DurationSec," _
        & "SourceServer,SourceDb,DestServer,DestDb,BranchID,FilterUsed," _
        & "SrcHeads,DstHeads,SrcDetails,DstDetails,SrcVATRows,DstVATRows," _
        & "SrcPayRows,DstPayRows,SrcPay2Rows,DstPay2Rows," _
        & "SrcAmount,DstAmount,SrcVATSum,DstVATSum,SrcTPaySum,DstTPaySum,SrcSPaySum,DstSPaySum," _
        & "BatchSize,FetchSize,IsOk,ErrorMsg) VALUES ("

    ' ‰’Ê’
    sql = sql & "'" & Esc(SessionCode, 32) & "',"
    sql = sql & "'" & Esc(direction, 20) & "',"
    sql = sql & "'" & Esc(TransferKind, 20) & "',"

    '  Ê«—ÌŒ + „œ…
    sql = sql & SQLDate(StartTime, True) & ","
    sql = sql & SQLDate(EndTime, True) & ","
    sql = sql & CStr(DurationSec) & ","

    ' „’«œ—/ÊÃÂ« 
    sql = sql & "'" & Esc(SourceServer, 128) & "',"
    sql = sql & "'" & Esc(SourceDb, 128) & "',"
    sql = sql & "'" & Esc(DestServer, 128) & "',"
    sql = sql & "'" & Esc(DestDb, 128) & "',"

    ' «·›—⁄ + «·›· —
    If BranchID = 0 Then
        sql = sql & "NULL,"
    Else
        sql = sql & CStr(BranchID) & ","
    End If
    sql = sql & "'" & Esc(FilterUsed, 2000) & "',"

    ' «·⁄œ¯«œ« 
    sql = sql & CStr(SrcHeads) & "," & CStr(DstHeads) & ","
    sql = sql & CStr(SrcDet) & "," & CStr(DstDet) & ","
    sql = sql & CStr(SrcVAT) & "," & CStr(DstVAT) & ","
    sql = sql & CStr(SrcPay) & "," & CStr(DstPay) & ","
    sql = sql & CStr(SrcPay2) & "," & CStr(DstPay2) & ","

    ' «·≈Ã„«·Ì« 
    sql = sql & FmtNum(SrcAmount) & "," & FmtNum(DstAmount) & ","
    sql = sql & FmtNum(SrcVATSum) & "," & FmtNum(DstVATSum) & ","
    sql = sql & FmtNum(SrcTPay) & "," & FmtNum(DstTPay) & ","
    sql = sql & FmtNum(SrcSPay) & "," & FmtNum(DstSPay) & ","

    ' ≈⁄œ«œ«  «· ‘€Ì· + «·‰ ÌÃ…
    sql = sql & CStr(BatchSize) & "," & CStr(FetchSize) & ","
    sql = sql & FmtBit(IsOk) & ","
    sql = sql & "'" & Esc(ErrorMsg, 4000) & "')"

    ' «ﬂ » ›Ì «·‰ﬁÿ…
    If Not cnPOS Is Nothing Then cnPOS.Execute sql, , adExecuteNoRecords
    ' «ﬂ » ›Ì «·”Ì—›—
    If Not cnSrv Is Nothing Then cnSrv.Execute sql, , adExecuteNoRecords
End Sub


'Private Function SqlStr(ByVal v As Variant) As String
'    If IsNull(v) Or Trim$(CStr(v)) = "" Then
'        SqlStr = "NULL"
'    Else
'        SqlStr = "'" & Replace(CStr(v), "'", "''") & "'"
'    End If
'End Function
'
'Private Function SqlNum(ByVal v As Variant) As String
'    If IsNull(v) Or Trim$(CStr(v)) = "" Then
'        SqlNum = "NULL"
'    Else
'        SqlNum = CStr(Val(v))
'    End If
'End Function
'
'Private Function SqlDateTime(ByVal v As Variant) As String
'    If IsNull(v) Or Trim$(CStr(v)) = "" Then
'        SqlDateTime = "NULL"
'    Else
'        '  ‰”Ìﬁ ISO ¬„‰
'        SqlDateTime = "'" & Format$(CDate(v), "yyyy-mm-dd hh:nn:ss") & "'"
'    End If
'End Function











Private Sub WriteTextLog(ByVal Msg As String)
    On Error Resume Next
    
    Dim f As Integer
    Dim LogFile As String
    
    LogFile = App.Path & "\LinkedServer_Errors.log"
    f = FreeFile
    
    Open LogFile For Append As #f
    Print #f, String(120, "-")
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & Msg
    Close #f
End Sub

Private Sub LogAdoErrors(ByVal CnX As ADODB.Connection, ByVal ProcName As String, ByVal SQLText As String, Optional ByVal ExtraInfo As String = "")
    On Error Resume Next
    
    Dim i As Long
    Dim s As String
    
    s = ""
    s = s & "Procedure: " & ProcName & vbCrLf
    
    If ExtraInfo <> "" Then
        s = s & "ExtraInfo: " & ExtraInfo & vbCrLf
    End If
    
    s = s & "SQL:" & vbCrLf & SQLText & vbCrLf
    
    If Not CnX Is Nothing Then
        s = s & "Connection State: " & CStr(CnX.State) & vbCrLf
        s = s & "Connection String: " & CnX.ConnectionString & vbCrLf
        
        If CnX.Errors.Count > 0 Then
            s = s & "ADO Errors Count: " & CStr(CnX.Errors.Count) & vbCrLf
            
            For i = 0 To CnX.Errors.Count - 1
                s = s & "ADO Error #" & CStr(i + 1) & vbCrLf
                s = s & "  Number      : " & CStr(CnX.Errors(i).Number) & vbCrLf
                s = s & "  Description : " & CnX.Errors(i).Description & vbCrLf
                s = s & "  Source      : " & CnX.Errors(i).Source & vbCrLf
                s = s & "  NativeError : " & CStr(CnX.Errors(i).NativeError) & vbCrLf
                s = s & "  SQLState    : " & CnX.Errors(i).SQLState & vbCrLf
            Next i
        Else
            s = s & "ADO Errors Count: 0" & vbCrLf
        End If
    Else
        s = s & "Connection object is Nothing" & vbCrLf
    End If
    
    s = s & "VB Err.Number      : " & CStr(Err.Number) & vbCrLf
    s = s & "VB Err.Description : " & Err.Description & vbCrLf
    s = s & "VB Err.Source      : " & Err.Source & vbCrLf
    
    WriteTextLog s
End Sub

Private Function BuildErrMsg(ByVal CnX As ADODB.Connection, ByVal ProcName As String, ByVal SQLText As String, Optional ByVal FriendlyPrefix As String = "") As String
    On Error Resume Next
    
    Dim s As String
    Dim i As Long
    
    s = ""
    
    If FriendlyPrefix <> "" Then
        s = FriendlyPrefix & vbCrLf & vbCrLf
    End If
    
    s = s & "Procedure: " & ProcName & vbCrLf
    
    If Err.Number <> 0 Then
        s = s & "VB Error Number: " & CStr(Err.Number) & vbCrLf
        s = s & "VB Error Description: " & Err.Description & vbCrLf
    End If
    
    If Not CnX Is Nothing Then
        If CnX.Errors.Count > 0 Then
            s = s & vbCrLf & "ADO Errors:" & vbCrLf
            
            For i = 0 To CnX.Errors.Count - 1
                s = s & "- Number: " & CStr(CnX.Errors(i).Number) & vbCrLf
                s = s & "  NativeError: " & CStr(CnX.Errors(i).NativeError) & vbCrLf
                s = s & "  Description: " & CnX.Errors(i).Description & vbCrLf
                s = s & "  Source: " & CnX.Errors(i).Source & vbCrLf
                s = s & "  SQLState: " & CnX.Errors(i).SQLState & vbCrLf
            Next i
        End If
    End If
    
    s = s & vbCrLf & "Last SQL:" & vbCrLf & SQLText
    
    BuildErrMsg = s
End Function

Private Function ExecSQL(ByVal CnX As ADODB.Connection, ByVal SQLText As String, ByVal ProcName As String, Optional ByVal ShowFriendlyName As String = "") As Boolean
    On Error GoTo EH
    
    mLastSQL = SQLText
    mLastProc = ProcName
    
    If CnX Is Nothing Then
        Err.Raise vbObjectError + 901, , "Connection object is Nothing"
    End If
    
    If CnX.State = adStateClosed Then
        Err.Raise vbObjectError + 902, , "Connection is closed"
    End If
    
    CnX.Errors.Clear
    CnX.Execute SQLText, , adCmdText
    
    ExecSQL = True
    Exit Function

EH:
    LogAdoErrors CnX, ProcName, SQLText, ShowFriendlyName
    ExecSQL = False
End Function

Private Sub SafeCloseRS(ByRef rs As ADODB.Recordset)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
    End If
    Set rs = Nothing
End Sub

Private Sub SafeCloseConn(ByRef CnX As ADODB.Connection)
    On Error Resume Next
    If Not CnX Is Nothing Then
        If CnX.State <> adStateClosed Then CnX.Close
    End If
    Set CnX = Nothing
End Sub

Private Function NzNum(ByVal v As Variant, Optional ByVal Def As Double = 0) As Double
    If IsNull(v) Or Trim$(v & "") = "" Then
        NzNum = Def
    Else
        NzNum = Val(v & "")
    End If
End Function

Private Function NzStr(ByVal v As Variant, Optional ByVal Def As String = "") As String
    If IsNull(v) Then
        NzStr = Def
    Else
        NzStr = v & ""
    End If
End Function


Private Function NormalizeDataSource(ByVal ServerName As String) As String
    Dim s As String
    s = Trim$(ServerName & "")
    
    If s = "" Then
        NormalizeDataSource = ""
        Exit Function
    End If
    
    If InStr(1, s, ",") > 0 Then
        NormalizeDataSource = s
    Else
        NormalizeDataSource = s & ",51433"
    End If
End Function

Private Function PrepareDebugTestConnections() As Boolean
    On Error GoTo errHandler

    TxtServerDataBaseName.Text = "Test"
    TxtPOSDB.Text = "Test"

    ServerDb = "Test"
    POSDb = "Test"

    PrepareDebugTestConnections = ConnectionFirst(False)
    Exit Function

errHandler:
    PrepareDebugTestConnections = False
End Function

Private Function GetDebugLogFileName(ByVal SessionCode As String) As String
    Dim p As String

    p = App.Path
    If Right$(p, 1) <> "\" Then p = p & "\"

    GetDebugLogFileName = p & "SyncDebug_" & SessionCode & ".log"
End Function

Private Sub DebugWriteLine(ByVal FileName As String, ByVal Msg As String)
    Dim ff As Integer

    On Error Resume Next
    ff = FreeFile
    Open FileName For Append As #ff
    Print #ff, Format$(Now, "dd/mm/yyyy hh:nn:ss AM/PM") & " - " & Msg
    Close #ff
End Sub

Private Sub DebugWriteSQL(ByVal FileName As String, ByVal Title As String, ByVal SQLText As String)
    Dim ff As Integer

    On Error Resume Next
    ff = FreeFile
    Open FileName For Append As #ff
    Print #ff, String$(90, "=")
    Print #ff, Format$(Now, "dd/mm/yyyy hh:nn:ss AM/PM") & " - " & Title
    Print #ff, "SQL:"
    Print #ff, SQLText
    Print #ff, String$(90, "=")
    Close #ff
End Sub
