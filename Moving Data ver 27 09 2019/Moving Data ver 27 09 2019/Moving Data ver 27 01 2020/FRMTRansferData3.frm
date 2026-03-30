VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
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
   Begin VB.CommandButton Command8 
      Caption         =   " ÕœÌÀ «·⁄„·«¡ „‰ «·„⁄—÷ «·Ï «·„’‰⁄"
      Height          =   495
      Left            =   120
      TabIndex        =   81
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtCountSalesOfeers 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3840
      TabIndex        =   79
      Top             =   2820
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
      Height          =   375
      Left            =   4560
      TabIndex        =   73
      Top             =   5640
      Width           =   1935
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
      Height          =   375
      Left            =   2130
      TabIndex        =   58
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   3810
      TabIndex        =   51
      Top             =   2070
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtCountSales 
      Height          =   375
      Left            =   3780
      TabIndex        =   50
      Top             =   1320
      Visible         =   0   'False
      Width           =   1755
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
      Text            =   "FRMTRansferData3.frx":000C
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
      Height          =   1275
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
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkSales 
         Caption         =   "«·„»Ì⁄« "
         Height          =   255
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
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
      Width           =   3375
      Begin VB.TextBox TxtServerDataBaseName 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Text            =   "byte"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox DestinationServer 
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   1815
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
      Width           =   3495
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
         Width           =   3345
      End
      Begin VB.TextBox TxtPOSDB 
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "LOCALPOS"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox POSlServer 
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DBname"
         Height          =   375
         Left            =   -240
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
         Left            =   120
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
      Width           =   1935
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
         Format          =   136314882
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
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   136577025
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
      FormatString    =   $"FRMTRansferData3.frx":0012
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
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   136642561
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker txtFromDate 
      Height          =   285
      Left            =   4860
      TabIndex        =   75
      Top             =   4860
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   136642561
      CurrentDate     =   41640
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ ⁄—Ê÷ «·”⁄— «·„‰ﬁÊ·…"
      Height          =   255
      Index           =   4
      Left            =   3900
      TabIndex        =   80
      Top             =   2520
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
      Top             =   390
      Width           =   4815
   End
   Begin VB.Label lblWait 
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
      Height          =   405
      Left            =   2250
      TabIndex        =   53
      Top             =   6150
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblCount 
      Height          =   315
      Left            =   5550
      TabIndex        =   52
      Top             =   6150
      Width           =   1425
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·„—œÊœ«  «·„‰ﬁÊ·…"
      Height          =   255
      Index           =   2
      Left            =   3870
      TabIndex        =   49
      Top             =   1770
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·›Ê« Ì— «·„‰ﬁÊ·…"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   48
      Top             =   1080
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
                
               mUserId = Val(Rs3!UserId & "")
               
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
         UnitId = IIf(IsNull(rsDouble_Entry("UnitId").Value), 0, rsDouble_Entry("UnitId").Value)
         ColorID = IIf(IsNull(rsDouble_Entry("ColorID").Value), 0, rsDouble_Entry("ColorID").Value)
         ItemSize = IIf(IsNull(rsDouble_Entry("ItemSize").Value), 0, rsDouble_Entry("ItemSize").Value)
         ClassId = IIf(IsNull(rsDouble_Entry("ClassId").Value), 0, rsDouble_Entry("ClassId").Value)
         
         
 
    sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transaction_Details]  (    "
sql = sql & "  Transaction_ID,  Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice,UnitId , ColorID, ItemSize, ClassId,SessionCode)"
 sql = sql & "   values (" & Transaction_ID & "," & Item_ID & ", " & ItemCase & "," & Quantity & "," & Price & "," & ItemDiscountType & "," & ItemDiscount & "," & ShowQty & "," & showPrice
 sql = sql & "," & UnitId & "," & ColorID & "," & ItemSize & "," & ClassId & "" & ",'" & SessionCode & "')"
 
           Cn.Execute sql
           rsDouble_Entry.MoveNext
    Next j
    
 
 
         
         
'ﬁÌœ «·”‰œ
  

sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,Transaction_ID,SessionCode)"
 sql = sql & " values( " & NoteId & ", " & SQLDate(Transaction_Date, True) & " , " & mNoteType & ", " & NoteSerial & ", " & NoteSerial1 & "," & BranchID & "," & Transaction_ID & ",'" & SessionCode & "')"
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
 POSConnection.Execute sql
 

  sql = "update   [" & POSDb & "].dbo.Transaction_Details" & "  set  Copied =1,SessionCode = '" & SessionCode & "' where Transaction_ID =" & FromTransaction_ID
 POSConnection.Execute sql

 
     StrSQL = "UPDATE  [" & ServerDb & "].dbo. Transactions SET NOTS=" & invoiceTransaction_ID & ",NOTS2= '" & invoiceNoteserial1 & "' ,SessionCode = '" & SessionCode & "' WHERE Transaction_ID=" & Transaction_ID
        Cn.Execute StrSQL
             StrSQL = "UPDATE  [" & ServerDb & "].dbo. Transactions SET NOTS=" & Transaction_ID & ",NOTS2= '" & NoteSerial1 & "',SessionCode = '" & SessionCode & "' WHERE Transaction_ID=" & invoiceTransaction_ID
        Cn.Execute StrSQL
        
        
End Function
Function ConnectionFirst(Optional ByVal IsLoad As Boolean = False) As Boolean

On Error GoTo ErrTrap
'«” ›”«—
'ServerDb = TxtServerDataBaseName.Text
'wael
'ServerDb = DestinationServer
' POSDb = TxtServerDataBaseName.Text


ServerDb = TxtServerDataBaseName.Text

     Set Cn = New ADODB.Connection
    With Cn
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & ServerDb & _
        ";Data Source=" & SysSQLServerName & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & ServerDb & _
                    ";Data Source=" & SysSQLServerName & ";Port=1433"
                    '";Data Source=" & ServerDb & ";Port=1433"
                    
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & ServerDb & ";Data Source=" & SysSQLServerName 'SysSQLServerName
                End If
          End If

.Open
End With
ConnectionFirst = True


'ServerDb = TxtServerDataBaseName.Text
'wael

If IsLoad Then Exit Function
POSDb = TxtPOSDB.Text
POSServer = POSlServer.Text


     Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & POSDb & _
        ";Data Source=" & POSServer & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & POSDb & _
                    ";Data Source=" & POSServer & ";Port=1433"
                    '";Data Source=" & ServerDb & ";Port=1433"
                    
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer 'SysSQLServerName
                End If
          End If

.Open

End With
ConnectionFirst = True

  

Dim mPosD  As String
Dim mServerD  As String
mPosD = "[" & POSlServer & "]" & ".Master.dbo."
mServerD = "[" & SysSQLServerName & "]" & ".Master.dbo."

Dim s As String
Dim ss As String
    
    s = " USE MASTER " & vbNewLine
    s = s & " DECLARE @sql NVARCHAR(4000) " & vbNewLine

    s = s & " DECLARE db_cursor CURSOR FOR " & vbNewLine
    s = s & "         select 'sp_dropserver ''' + [srvName] + '''' from sysservers " & vbNewLine

    s = s & "     OPEN db_cursor " & vbNewLine
    s = s & "     FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine

    s = s & "     WHILE @@FETCH_STATUS = 0 " & vbNewLine
    s = s & "     BEGIN " & vbNewLine

    s = s & "            EXEC (@sql) " & vbNewLine

    s = s & "            FETCH NEXT FROM db_cursor INTO @sql " & vbNewLine
    s = s & "     End " & vbNewLine

    s = s & "     Close db_cursor " & vbNewLine
    s = s & "     DEALLOCATE db_cursor " & vbNewLine
    
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute s & ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute s & ss
   
Dim rsDummy As New ADODB.Recordset
's = "select * from " & mServerD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
'rsDummy.Open s, Cn, adOpenStatic
'If rsDummy.EOF Then
'    Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If
'rsDummy.Close

's = "select * from sys.servers Where name Like '" & SysSQLServerName & "'"


's = "select * from sys.servers Where name Like '" & POSServer & "'"
s = "select * from sysservers Where srvName Like '" & POSServer & "'"
rsDummy.Open s, Cn, adOpenStatic
If rsDummy.EOF Then
    Cn.Execute "EXEC sp_addlinkedserver [" & POSServer & "]"
   ' Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If
  


's = "select * from " & mServerD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
s = "select * from sysservers Where srvName Like '" & SysSQLServerName & "'"
rsDummy.Close
rsDummy.Open s, Cn, adOpenStatic
If rsDummy.EOF Then
   
    Cn.Execute "EXEC sp_addlinkedserver [" & SysSQLServerName & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If


'rsDummy.Close
s = " Use Master "
POSConnection.Execute s

's = "select * from " & mPosD & "sysservers Where srvName Like '" & SysSQLServerName & "'"
s = "select * from sysservers Where srvName Like '" & SysSQLServerName & "'"
rsDummy.Close
rsDummy.Open s, POSConnection, adOpenStatic
If rsDummy.EOF Then
    POSConnection.Execute " EXEC sp_addlinkedserver [" & SysSQLServerName & "]"

   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If

rsDummy.Close

s = "select * from sysservers Where srvName Like '" & POSServer & "'"

rsDummy.Open s, POSConnection, adOpenStatic
If rsDummy.EOF Then
    
    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If
rsDummy.Close



s = "select * from " & mPosD & "sysservers Where srvName Like '" & POSServer & "'"
rsDummy.Open s, POSConnection, adOpenStatic
If rsDummy.EOF Then

    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
End If
rsDummy.Close



Set rsDummy = New ADODB.Recordset
s = "Select * from [" & SysSQLServerName & "]." & ServerDb & ".dbo.TblOptions "
rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
If Not rsDummy.EOF Then
    NoOFDigitUserTrans = Val(rsDummy!NoOFDigitUserTrans & "")
    StoreDigit = Val(rsDummy!StoreDigit & "")
    BranchDigit = Val(rsDummy!BranchDigit & "")
    IsSerialByUserTrans = Val(rsDummy!IsSerialByUserTrans & "")
    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
    InstallmntsvchrCoding = Val(rsDummy!InstallmntsvchrCoding & "")
    ExpensesCoding2 = Val(rsDummy!ExpensesCoding2 & "")
    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
    ExpensesCoding = Val(rsDummy!ExpensesCoding & "")
    AllowProjectBill2Serial = Val(rsDummy!AllowProjectBill2Serial & "")
    NoOFDigitUserVouc = Val(rsDummy!NoOFDigitUserVouc & "")
    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
    IsSerialByUserVouch = Val(rsDummy!IsSerialByUserVouch & "")
    JLCodeBasedOnBranch = Val(rsDummy!JLCodeBasedOnBranch & "")
    
End If

rsDummy.Close
'
's = "select * from sys.servers Where name Like '" & POSServer & "'"
'rsDummy.Open s, POSConnection, adOpenStatic
'If rsDummy.EOF Then
'    POSConnection.Execute " EXEC sp_addlinkedserver [" & POSServer & "]"
'   ' Cn.Execute " EXEC sp_addlinkedsrvlogin '#" & POSServer & "#', 'false', NULL, '#username#', '#password@123" '"
'End If



'Do While Not rsDummy.EOF
'
'
'    rsDummy.MoveNext
'Loop



Exit Function
ErrTrap:
Text1 = Cn.ConnectionString
Text2 = POSConnection.ConnectionString
MsgBox "Õÿ√ ›Ì «·« ’«·"
 ConnectionFirst = False


End Function

Private Function DeleteLinkedServer()
 

    
    
End Function

Private Sub cmdUdateFiles_Click()



On Error GoTo EE:
'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

Command4_Click
lblWait.Visible = True
   
   
  UpdateFiles POSlServer, POSDb, "TblBranchesData", "branch_id"
  UpdateFiles POSlServer, POSDb, "cachierData", "Id"
  UpdateFiles POSlServer, POSDb, "TblStore", "StoreID"
  UpdateFiles POSlServer, POSDb, "TblBoxesData", "BoxId"
  UpdateFiles POSlServer, POSDb, "BanksData", "BankId"
  UpdateFiles POSlServer, POSDb, "TblEmployee", "Emp_ID"
  
'  UpdateFiles POSlServer, POSDb, "ACCOUNTS", "Account_ID"
  UpdateFiles POSlServer, POSDb, "TblCustemers", "CusId"
  
   Exit Sub
EE:
MsgBox "BasicData"
End Sub


Private Sub cmdUpdatePrice_Click()
Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
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
             mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
             mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
             mServerD = ServerDb & ".dbo."
            
 
            POSConnection.Execute "Delete " & mPosD & "TblItemsUnits "
 
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
        
    

    
    
                  
             MsgBox " „ ‰ﬁ· »Ì«‰«  «·ÊÕœ« "
         
    
            
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
            
        
            
            

            
            
             MsgBox " „ ‰ﬁ· »Ì«‰«  «·«’‰«›"
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
    
        If Val(rsDummy!UserId & "") < 10 Then
            InUser = "0" + CStr(Val(rsDummy!UserId & ""))
        
        
        Else
        
            InUser = CStr(Val(rsDummy!UserId & ""))
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

   MsgBox " „ ÷»ÿ «·”Ì—Ì«·"
 

End Sub

Private Sub Command1_Click()
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
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
            "Persist Security Info=False;Initial Catalog=" & POSDb & _
            ";Data Source=" & POSlServer & ";Port=1433"
        
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
 
                    FromUserID = IIf(IsNull(Rs3("UserID").Value), 0, Rs3("UserID").Value)
                    FromEmp_ID = IIf(IsNull(Rs3("Emp_ID").Value), 0, Rs3("Emp_ID").Value)
                    FromStoreID = IIf(IsNull(Rs3("storeID").Value), 0, Rs3("storeID").Value)
                    
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
                                FromCusID = Val(rsCus2!cusId & "")
                            End If
                            
                        End If
                        
                        
                    Else
                        FromCusID = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                    End If
                    
                    FromBoxid = IIf(IsNull(Rs3("Boxid").Value), 0, Rs3("Boxid").Value)
                    POSBillType = IIf(IsNull(Rs3("POSBillType").Value), 0, Rs3("POSBillType").Value)
                    FromUserID = Val(Rs3!UserId & "")
                    mTransaction_NetValue = Val(Rs3!Transaction_NetValue & "")
                    FromPaymentType = IIf(IsNull(Rs3("PaymentType").Value), 0, Rs3("PaymentType").Value)
                    FromBillBasedOn = IIf(IsNull(Rs3("BillBasedOn").Value), 0, Rs3("BillBasedOn").Value)
                    FromVATYou = IIf(IsNull(Rs3("VATYou").Value), 0, Rs3("VATYou").Value)
                    FromVAT = IIf(IsNull(Rs3("VAT").Value), 0, Rs3("VAT").Value)
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
                    FromNoteId = IIf(IsNull(Rs3("NoteId").Value), 0, Rs3("NoteId").Value) ' —ﬁ„ ﬁÌœ «·›« Ê—…
                    
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
                    Transaction_Serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=" & Transaction_Type & ""))
                    NoteSerial1 = Voucher_coding(FromBranchID, FromTransaction_Date, mSanadNo, mNoteType, , Transaction_Type, , , , , , FromUserID)
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
                    sql = sql & "  Transaction_ID,Transaction_Date,"
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
                    sql = sql & "CashCustomerPhone,last_changed ,NetValue,Transaction_NetValue,DepandToConv  )"
                    
                    
                    
                    
                    sql = sql & "   values (" & Transaction_ID & "," & SQLDate(Transaction_Date, True) & ","
                    sql = sql & Transaction_Serial & ","
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
                    sql = sql & "'" & CashCustomerPhone & "' ," & SQLDate(last_changed, True) & "," & NetValue & "," & mTransaction_NetValue & "," & IIf(DepandToConv, 1, 0) & " )"
                    
                    
       



                    
                    
                    '   fromTransaction_Serial
                    Text1.Text = sql
                   ' Exit Sub
                   Text4 = ""
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
                        UnitId = IIf(IsNull(rsDouble_Entry("UnitId").Value), 0, rsDouble_Entry("UnitId").Value)
                        ColorID = IIf(IsNull(rsDouble_Entry("ColorID").Value), 0, rsDouble_Entry("ColorID").Value)
                        ItemSize = IIf(IsNull(rsDouble_Entry("ItemSize").Value), 0, rsDouble_Entry("ItemSize").Value)
                        ClassId = IIf(IsNull(rsDouble_Entry("ClassId").Value), 0, rsDouble_Entry("ClassId").Value)
                        mmVatyo = IIf(IsNull(rsDouble_Entry("Vatyo").Value), 0, rsDouble_Entry("Vatyo").Value)
                    
     
                        sql = " INSERT INTO  [" & ServerDb & "].[dbo].[Transaction_Details]  (    "
                        sql = sql & "  Transaction_ID,  Item_ID, ItemCase, Quantity, Price, ItemDiscountType, ItemDiscount, ShowQty, showPrice,UnitId , ColorID, ItemSize, ClassId,SessionCode,Vatyo)"
                        sql = sql & "   values (" & Transaction_ID & "," & Item_ID & ", " & ItemCase & "," & Quantity & "," & Price & "," & ItemDiscountType & "," & ItemDiscount & "," & ShowQty & "," & showPrice
                        sql = sql & "," & UnitId & "," & ColorID & "," & ItemSize & "," & ClassId & "" & ",'" & SessionCode & "'," & mmVatyo & ")"
                        
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
                        
                        
                        Cn.Execute sql
                        rsDouble_Entry.MoveNext
                    Next j
  ' MsgBox "4"
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
                            boxid = IIf(IsNull(rsDouble_Entry("boxid").Value), 0, rsDouble_Entry("boxid").Value)
                            Recorddate = IIf(IsNull(rsDouble_Entry("Recorddate").Value), 0, rsDouble_Entry("Recorddate").Value)
                            PointID = IIf(IsNull(rsDouble_Entry("PointID").Value), 0, rsDouble_Entry("PointID").Value)
                            CurrentCashireID = IIf(IsNull(rsDouble_Entry("CurrentCashireID").Value), 0, rsDouble_Entry("CurrentCashireID").Value)
                            PaymentID = IIf(IsNull(rsDouble_Entry("PaymentID").Value), 0, rsDouble_Entry("PaymentID").Value)
                            Value = IIf(IsNull(rsDouble_Entry("Value").Value), 0, rsDouble_Entry("Value").Value)
                            CardNo = IIf(IsNull(rsDouble_Entry("CardNo").Value), 0, rsDouble_Entry("CardNo").Value)
                            Effect = IIf(IsNull(rsDouble_Entry("Effect").Value), 0, rsDouble_Entry("Effect").Value)
                            
                            
                            sql = " INSERT INTO  [" & ServerDb & "].[dbo].[TblTransactionPayments]  (    "
                            sql = sql & "  Transaction_ID,  boxid, Recorddate, PointID, CurrentCashireID, PaymentID,Value,CardNo,Effect,SessionCode)"
                            sql = sql & "   values (" & Transaction_ID & "," & boxid & ", " & SQLDate(Recorddate, True) & "," & PointID & "," & CurrentCashireID & "," & PaymentID & "," & Value & ",'" & CardNo & "'," & Effect & ",'" & SessionCode & "')"
                            
                            
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
                cusId = IIf(IsNull(Rs3("CusID").Value), 0, Rs3("CusID").Value)
                NoteCashingType = IIf(IsNull(Rs3("NoteCashingType").Value), 0, Rs3("NoteCashingType").Value)
                boxid = IIf(IsNull(Rs3("BoxID").Value), "Null", Rs3("BoxID").Value)
                ChqueNum = IIf(IsNull(Rs3("ChqueNum").Value), 0, Rs3("ChqueNum").Value)
                DueDate = IIf(IsNull(Rs3("DueDate").Value), 0, Rs3("DueDate").Value)
                ChequeBoxID = IIf(IsNull(Rs3("ChequeBoxID").Value), 0, Rs3("ChequeBoxID").Value)
                BankID = IIf(IsNull(Rs3("BankID").Value), "Null", Rs3("BankID").Value)
                TotalNotesValue = IIf(IsNull(Rs3("TotalNotesValue").Value), 0, Rs3("TotalNotesValue").Value)
                
                sql = " INSERT INTO [" & ServerDb & "].[dbo].[Notes]([NoteID], [NoteDate], [NoteType], [NoteSerial], [NoteSerial1] ,branch_no,UserID,CashingType,EmpId,VAT"
                 sql = sql & ",NCashingType, Status,Note_Value,BankName,Remark,CusID,NoteCashingType,BoxID,ChqueNum,DueDate,ChequeBoxID,BankID,TotalNotesValue,copied,SessionCode )"
                 sql = sql & " values( " & NoteId & ", " & SQLDate(CDate(NoteDate), True) & " , 4, " & NoteSerial & ", " & NoteSerial1 & "," & BranchID & ",1," & CashingType & "," & EmpId & "," & VAT
                 sql = sql & "," & NCashingType & ", " & Status & "," & Note_Value & ",'" & BankName & "','" & Remark & "'," & cusId & "," & NoteCashingType & "," & boxid & "," & ChqueNum & "," & SQLDate(CDate(Date), True) & "," & ChequeBoxID & "," & BankID & "," & TotalNotesValue & ",1,'" & SessionCode & "')"
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

MsgBox " „ ‰ﬁ· «·»Ì«‰« "

 
    





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

Private Sub Command2_Click()

'   ************************************'check items here first wael*******************
 Dim StrSQL As String
If POSlServer.Text = "" Then
MsgBox "«Œ — «·‰ﬁÿÂ «·„‰ﬁÊ· „‰Â« «Ê·«", vbCritical, "OFFLINE"
Exit Sub
End If

'Command4_Click
lblWait.Visible = True
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
    
    sql = " select count (ItemID ) As NoOfitems ,max(ItemID) as MaxItemid from TblItems  "
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
    Rs3.Close
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' MsgBox "Step 1"
    If Rs3.RecordCount > 0 Then
        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close
    
   ' MsgBox "Item Server" & NoOFItem_Server
   ' MsgBox "Item Pos" & NoOFItem_POS
    'step 2
   ' Exit Sub
    If NoOFItem_Server <> NoOFItem_POS Then
             'checkGroup
        Dim NoOfGroups_pos As Double
        Dim NoOfGroups_server As Double
             
        Dim MaxGroupid_pos As Double
        Dim MaxGroupidserver As Double
        
                       
        sql = " select count (GroupID ) As NoOfGroups ,max(GroupID) as MaxGroupid from Groups  "
        Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
        If Rs3.RecordCount > 0 Then
            NoOfGroups_pos = IIf(IsNull(Rs3("NoOfGroups").Value), 0, Rs3("NoOfGroups").Value)
            MaxGroupid_pos = IIf(IsNull(Rs3("MaxGroupid").Value), 0, Rs3("MaxGroupid").Value)
        End If
        Rs3.Close
        'MsgBox "Step 2"
             
             
        sql = " select count (GroupID ) As NoOfGroups ,max(GroupID) as MaxGroupid from Groups  "
 
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs3.RecordCount > 0 Then
            NoOfGroups_server = IIf(IsNull(Rs3("NoOfGroups").Value), 0, Rs3("NoOfGroups").Value)
            MaxGroupidserver = IIf(IsNull(Rs3("MaxGroupid").Value), 0, Rs3("MaxGroupid").Value)
        End If
        Rs3.Close
        
         'MsgBox "Step 3"
        Dim s As String
        
        If NoOFItem_Server <> NoOFItem_POS Then
            
   
            BolFrmLoaded = True
    
              
         Do While Rs3.State = adStateExecuting
                DoEvents
            Loop
     
              
            s = ""
             
            Dim mPosD As String
            Dim mServerD As String
            mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
            mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
            mServerD = ServerDb & ".dbo."
            
            s = " INSERT INTO " & mPosD & "Groups"
            s = s & " SELECT *"
            s = s & " FROM   " & mServerD & "Groups T2"
            s = s & " WHERE  T2.GroupID NOT IN (SELECT IsNull(Tpos.GroupID,0)"
            s = s & "                                      FROM   " & mPosD & "Groups  as Tpos);"
            
           ' Text4 = s
           ' Exit Sub
            Cn.Execute s
           ' MsgBox "Step 4"
            s = " INSERT INTO " & mPosD & "TblUnites"
            s = s & " SELECT *"
            s = s & " FROM   " & mServerD & "TblUnites T2"
            s = s & " WHERE  T2.UnitID NOT IN (SELECT IsNull(Tpos.UnitID,0)"
            s = s & "                                      FROM   " & mPosD & "TblUnites as  Tpos);"
            
            Cn.Execute s
            
           ' MsgBox "Step 5"
'            s = " INSERT INTO " & mPosD & "TblItemLoc"
'            s = s & " SELECT *"
'            s = s & " FROM   " & mServerD & "TblItemLoc T2"
'            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'            s = s & "                                      FROM   " & mPosD & "TblItemLoc);"
'
'            Cn.Execute s
'

            
            s = " INSERT INTO " & mPosD & "TblItems"
            s = s & " SELECT * "
            s = s & " FROM   " & mServerD & "TblItems T2"
            s = s & " WHERE  T2.ItemID NOT IN (SELECT IsNull(Tpos.ItemID,0)"
            s = s & "                                      FROM   " & mPosD & "TblItems as Tpos);"
            
            
            Cn.Execute s
            
            
           ' MsgBox "Step 6"
           '1
'            MsgBox "pos" & mPosD
'            MsgBox "Server " & mServerD
'
'                       s = " INSERT INTO " & mPosD & "TblItemProductLine"
'            s = s & " SELECT *"
'            s = s & " FROM   " & mServerD & "TblItemProductLine T2"
'            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'            s = s & "                                      FROM   " & mPosD & "TblItemProductLine);"
'
'            Cn.Execute s
'
            
'            s = " INSERT INTO " & mPosD & "TblItemsAttach"
'            s = s & " SELECT *"
'            s = s & " FROM   " & mServerD & "TblItemsAttach T2"
'            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'            s = s & "                                      FROM   " & mPosD & "TblItemsAttach);"
'
'            Cn.Execute s
'
'            s = " INSERT INTO " & mPosD & "ItemsPrice"
'            s = s & " SELECT *"
'            s = s & " FROM   " & mServerD & "ItemsPrice T2"
'            s = s & " WHERE  T2.Item_ID NOT IN (SELECT Item_ID"
'            s = s & "                                      FROM   " & mPosD & "ItemsPrice);"
'
'            Cn.Execute s
'
'
            
'            s = " INSERT INTO  " & mPosD & "ItemsParts"
'            s = s & " SELECT *"
'            s = s & " FROM   " & mServerD & "ItemsParts T2"
'            s = s & " WHERE  T2.ItemID NOT IN (SELECT ItemID"
'            s = s & "                                      FROM   " & mPosD & "ItemsParts);"
'
'            Cn.Execute s
            
        '    MsgBox "Step 6"
            s = " INSERT INTO " & mPosD & "TblItemsUnits"
            s = s & " SELECT *"
            s = s & " FROM   " & mServerD & "TblItemsUnits T2"
            s = s & " WHERE  T2.ItemID NOT IN (SELECT IsNull(TPos.ItemID,0)"
            s = s & "                                      FROM   " & mPosD & "TblItemsUnits as TPos);"
                                     
            Cn.Execute s
            
            Text5 = s
            
     '       MsgBox "Step 7"

             
            'Copy  remains Groups
            'Copy  remains Items
            'Copy itemsunits
            
            
             MsgBox " „ ‰ﬁ· »Ì«‰«  «·«’‰«›"
             Command2.Enabled = False
    
        End If
     Else
      MsgBox "    „·›   «·«’‰«› „ÕœÀ"
      lblWait.Visible = False
 
End If
    
    
   '************************************'check items here first*******************

End Sub

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
    MsgBox "⁄œœ «’‰«› «·‰ﬁÿ…" & NoOFItem_POS
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
    
    MsgBox " „ «·‰ﬁ·"
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

   MsgBox " „ ÷»ÿ «·”Ì—Ì«·"
 
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
    
    
    
  «·«”
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
    rsDummy2!cusId = mMaxId
  
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
MsgBox "BasicData"

End Sub

Private Sub DCboBranch_Click(Area As Integer)
txtBranch = DCboBranch.BoundText
End Sub

Private Sub Form_Load()
'21 11 2017
' „  ‰›Ì– «·„»Ì⁄«  ﬂ«„·… „⁄ ﬁÌœÂ«  „⁄ ”‰œ «·’—› „⁄ ﬁÌœ…
'
'
'
 On Error Resume Next
txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
TxtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
TxtUSERID = GetSetting("ConvertToAccess", "Setting", "USERID", "USERID")
TxtCHECKTIME = GetSetting("ConvertToAccess", "Setting", "CHECKTIME", "CHECKTIME")
'DcTime.Value = GetSetting("ConvertToAccess", "Setting", "UpdateHours", "00")
dbRecordDate = Date
TxtServerDataBaseName = SysSQLServerDataBaseName
DestinationServer = SysSQLServerName

txtFromDate.Value = Date
txtToDate.Value = Date
'BranchDigit = 1
Dim Msg As String
If Dir(App.Path & "\pos.txt", vbNormal) = "" Then
            Msg = "„·›  ”ÃÌ· «·ﬁÊ«⁄œ €Ì— „ÊÃÊœ ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
           End
           
        End If
        
    Open App.Path & "\pos.txt" For Input As #1
    POSname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
             POSname.AddItem (VarSet(0))
                ServersName.AddItem (VarSet(1))
            DbName.AddItem (VarSet(2))
                            
            End If
        End If
    
    Loop
   Dim StrSQL As String

    
 If ConnectionFirst(True) = False Then
        Exit Sub
    End If

        StrSQL = "SELECT branch_id,branch_name FROM TblBranchesData"
 




    GetComboData DCboBranch, StrSQL
    
    Close #1


 
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
        .BoundColumn = rs(0).name
        .ListField = rs(1).name

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
    dbRecordDate = grd.TextMatrix(grd.Row, grd.ColIndex("Transaction_Date"))
End If
End Sub

Private Sub POSname_Change()
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
    
    sql = " "
    
    sql = sql & "     SELECT SUM(CountSales) CountSales ,SUM(CountReturn) CountReturn,SUM(CountSalesOfeers) CountSalesOfeers,Transaction_Date FROM ("
    sql = sql & "         SELECT COUNT(t.Transaction_ID)     CountTotal,"
    sql = sql & "                CountSales       = ("
    sql = sql & "                    Case t.Transaction_Type"
    sql = sql & "                         WHEN 21 THEN COUNT(t.Transaction_ID)"
    sql = sql & "                         ELSE 0"
    sql = sql & "                    End"
    sql = sql & "                ),"
    sql = sql & "                CountReturn     = ("
    sql = sql & "                    Case t.Transaction_Type"
    sql = sql & "                         WHEN 9 THEN COUNT(t.Transaction_ID)"
    sql = sql & "                         ELSE 0"
    sql = sql & "                    End"
    sql = sql & "                ),"
   
    sql = sql & "                CountSalesOfeers     = ("
    sql = sql & "                    Case t.Transaction_Type"
    sql = sql & "                         WHEN 42 THEN COUNT(t.Transaction_ID)"
    sql = sql & "                         ELSE 0"
    sql = sql & "                    End"
    sql = sql & "                ),"
    
    sql = sql & "                t.Transaction_Date,"
    sql = sql & "                Transaction_Type"
    sql = sql & "         FROM   Transactions             AS t"
    sql = sql & "         Where IsNull(t.Copied, 0) = 0"
    sql = sql & "                AND (t.Transaction_Type = 9 OR t.Transaction_Type = 21 OR t.Transaction_Type = 42)"
    sql = sql & "         Group By"
    sql = sql & "                Transaction_Date,"
    sql = sql & "                Transaction_Type"
        
    sql = sql & "         ) T"
    sql = sql & "         Group By"
    sql = sql & "                Transaction_Date"
    sql = sql & "         Order By"
    sql = sql & "                Transaction_Date"

     Text5 = sql
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    grd.Rows = 1
    grd.Rows = 2
    Do While Not Rs3.EOF
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountSales")) = Rs3!CountSales & ""
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountReturn")) = Rs3!CountReturn & ""
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("CountSalesOfeers")) = Rs3!CountSalesOfeers & ""
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("Transaction_Date")) = Rs3!Transaction_Date & ""
        Rs3.MoveNext
        grd.Rows = grd.Rows + 1
    Loop
    Rs3.Close
End Sub

Private Sub POSname_Click()
On Error Resume Next
    DbName.ListIndex = POSname.ListIndex
    ServersName.ListIndex = POSname.ListIndex
     
   POSlServer.Text = ServersName.Text
    TxtPOSDB.Text = DbName.Text
    
    POSname_Change
    
    
    
End Sub
Private Function GetQuery() As String
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
    If chkSales.Value = vbChecked Then
        tempString = tempString & "," & 21
    End If
    If chkSalesReturn.Value = vbChecked Then
        tempString = tempString & "," & 9
        
    End If
    
    If chkSalesOffers.Value = vbChecked Then
        tempString = tempString & "," & 42
        
        
    End If
    
    
    'GetTransIds = tempString
    
    s = s & "  (Transaction_Type in ( " & tempString & " ) )"
     's = s and dbo.Transactions.Transaction_Date ='" & SQLDate(dbRecordDate.Value, False) & "')"
     
     s = s & " and (dbo.Transactions.Transaction_Date >='" & SQLDate(txtFromDate.Value, False) & "')"
     s = s & " and (dbo.Transactions.Transaction_Date <='" & SQLDate(txtToDate.Value, False) & "')"
    
     If chkSalesOffers.Value = vbChecked Then
        s = s & " and  (IsNull(Transactions.DepandToConv,0) = 1  Or Transactions.Transaction_ID In (Select ApprovalData.Transaction_ID from ApprovalData "
        s = s & " where IsNull(ApprovDate,'') <> '' and ScreenName = 'FrmPO1' ))"
        
     End If
    
GetQuery = s
End Function

Private Sub VSFlexGrid1_Click()

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

