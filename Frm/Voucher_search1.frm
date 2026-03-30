VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Voucher_search1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ”‰œ ’—› „ ⁄œœ"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13830
   Icon            =   "Voucher_search1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   13830
   Begin VB.CheckBox Check6 
      Alignment       =   1  'Right Justify
      Caption         =   "·›—⁄ „Õœœ"
      Height          =   255
      Left            =   12720
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   7
      Left            =   120
      TabIndex        =   58
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame FraNote 
      Height          =   1725
      Left            =   165
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   1560
      Width           =   6255
      Begin VB.TextBox TxtChequeNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   450
         Width           =   4365
      End
      Begin VB.CheckBox Check7 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ì"
         Height          =   195
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   900
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰"
         Height          =   195
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox txtperson 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1200
         Width           =   4365
      End
      Begin MSComCtl2.DTPicker DtpChequeDueDatefrom 
         Height          =   315
         Left            =   2280
         TabIndex        =   50
         Top             =   840
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104857601
         CurrentDate     =   39614
      End
      Begin MSDataListLib.DataCombo DcboBankName 
         Height          =   315
         Left            =   30
         TabIndex        =   51
         Top             =   120
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpChequeDueDateTo 
         Height          =   315
         Left            =   30
         TabIndex        =   67
         Top             =   840
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104857601
         CurrentDate     =   39614
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„” ›Ìœ"
         Height          =   285
         Index           =   34
         Left            =   4620
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1335
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   285
         Index           =   16
         Left            =   4620
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   285
         Index           =   17
         Left            =   4620
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   450
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
         Height          =   285
         Index           =   18
         Left            =   4620
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   780
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Height          =   3495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   4200
      Width           =   13815
      Begin VSFlex8UCtl.VSFlexGrid DataGrid2 
         Height          =   3165
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   13575
         _cx             =   23945
         _cy             =   5583
         Appearance      =   1
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
         BackColorFixed  =   -2147483633
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   150
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"Voucher_search1.frx":000C
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   6
      Left            =   8520
      TabIndex        =   45
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   11160
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtAccCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   5280
      TabIndex        =   43
      Top             =   -3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   -1560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊŸ›"
         Height          =   195
         Index           =   3
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ê—œ"
         Height          =   195
         Index           =   2
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„Ì·"
         Height          =   195
         Index           =   1
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»"
         Height          =   195
         Index           =   0
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " ⁄œÌ· «·‘—Õ ·”ÿ— „⁄Ì‰"
      Height          =   1695
      Left            =   14760
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   6600
      Width           =   9015
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "DEV_DES"
         DataSource      =   "Adodc1"
         Height          =   1215
         Left            =   1080
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   240
         Width           =   7815
      End
      Begin ALLButtonS.ALLButton ALLButton5 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Voucher_search1.frx":0196
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«” »œ«· ﬂ·„…"
      Height          =   1455
      Left            =   14160
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   5520
      Width           =   3375
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "‰›–"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Voucher_search1.frx":01B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txt2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«” »œ«· »"
         Height          =   255
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "«·»ÕÀ ⁄‰"
         Height          =   255
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Õœœ «· «—ÌŒ"
      Height          =   615
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3480
      Width           =   4455
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ï"
         Height          =   195
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰"
         Height          =   195
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTP_Date 
         Height          =   330
         Left            =   2160
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   104857603
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Date1 
         Height          =   330
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   104857603
         CurrentDate     =   37140
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Voucher_search1.frx":01CE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â «·‘«‘…"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Voucher_search1.frx":01EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   11160
      TabIndex        =   19
      Tag             =   "Press F3 To Search"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   -1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   4
      Left            =   6720
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   " Õ ÊÏ «·ﬂ·„…"
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   -1440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "«·ﬂ·„… ›ﬁÿ"
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   -1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   6720
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   5880
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   8400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "„Ê«›ﬁ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Voucher_search1.frx":0206
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   585
      Left            =   1080
      Top             =   -1680
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1032
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
      Caption         =   " Õ—Ìﬂ"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   2520
      Top             =   9000
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1032
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
      Caption         =   " Õ—Ìﬂ"
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
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   8400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â ﬂ‘› Õ”«»"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Voucher_search1.frx":0222
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Bindings        =   "Voucher_search1.frx":023E
      Height          =   315
      Left            =   5880
      TabIndex        =   61
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   16777215
      ListField       =   "account_name"
      BoundColumn     =   "code"
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ALLButtonS.ALLButton ALLButton6 
      Height          =   495
      Left            =   4080
      TabIndex        =   62
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "»ÕÀ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Voucher_search1.frx":0253
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DCExtraAccount 
      Height          =   315
      Left            =   6720
      TabIndex        =   63
      Top             =   3000
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ALLButtonS.ALLButton ALLButton7 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   65
      Top             =   3600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "„”Õ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Voucher_search1.frx":026F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ì „»·€"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   59
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·”‰œ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   -1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12600
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„’œ—"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   -1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰ „»·€"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·»ÕÀ ⁄‰ ”‰œ ’—› „ ⁄œœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   705
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "‰ÿ«ﬁ «·»ÕÀ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12240
      TabIndex        =   13
      Top             =   -1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ «·⁄«„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label title_lbl 
      Caption         =   "Label8"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label case_id 
      Caption         =   "0"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   -720
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Voucher_search1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim first_run As Boolean

Private Sub ALLButton2_Click()
    On Error Resume Next
     
    If Check3.value = vbChecked And Check4.value = vbChecked Then
        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%'  and recorddate >= CONVERT(DATETIME, '" & DTP_Date.year & "-" & DTP_Date.Month & "-" & DTP_Date.day & " 00:00:00', 102) and recorddate <= CONVERT(DATETIME, '" & DTP_Date1.year & "-" & DTP_Date1.Month & "-" & DTP_Date1.day & " 00:00:00', 102)"
    ElseIf Check3.value = vbChecked And Check4.value = Unchecked Then
     
        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%'  and recorddate >= CONVERT(DATETIME, '" & DTP_Date.year & "-" & DTP_Date.Month & "-" & DTP_Date.day & " 00:00:00', 102) "
    ElseIf Check3.value = Unchecked And Check4.value = vbChecked Then
        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%' and  recorddate <= CONVERT(DATETIME, '" & DTP_Date1.year & "-" & DTP_Date1.Month & "-" & DTP_Date1.day & " 00:00:00', 102)"
    Else
        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%' "
      
    End If
   
    'Sql = "select * from RptLedger_Sub where     recorddate = CONVERT(DATETIME, '" & DTP_Date.year & "-" & DTP_Date.Month & "-" & DTP_Date.Day & " 00:00:00', 102)"
    Retrive sql
End Sub
Private Sub Retrive(Optional sql As String)
    Dim Num As Integer
    On Error GoTo ErrTrap
     Dim rs As ADODB.Recordset
          If Combo1.ListIndex = 0 Then
       ' sql = sql & " and  NoteType=53 "
       '
    ElseIf Combo1.ListIndex = 1 Then
       ' sql = sql & " and   NoteType<>53 "
     
    End If
 
    DataGrid2.Clear flexClearScrollable, flexClearEverything
    Set rs = New ADODB.Recordset
     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

If rs.RecordCount > 0 Then

           If Not (rs.EOF Or rs.BOF) Then
        DataGrid2.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With DataGrid2
                .TextMatrix(Num, .ColIndex("Notes_ID")) = IIf(IsNull(rs("Notes_ID").value), "", rs("Notes_ID").value)
               .TextMatrix(Num, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
            
                    .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", Trim(rs("NoteSerial1").value))
                     If case_id.Caption = 2 Then
                     .ColHidden(.ColIndex("ManualNo")) = False
                    .TextMatrix(Num, .ColIndex("ManualNo")) = IIf(IsNull(rs("ManualNo").value), "", Trim(rs("ManualNo").value))
                    Else
                    .ColHidden(.ColIndex("ManualNo")) = True
                   End If
                    .TextMatrix(Num, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", Trim(rs("Account_Serial").value))
            

               If SystemOptions.UserInterface = ArabicInterface Then
                If val(rs("Credit_Or_Debit").value) = 0 Then
            .TextMatrix(Num, .ColIndex("Credit_Or_Debit")) = "„œÌ‰"
           Else
            .TextMatrix(Num, .ColIndex("Credit_Or_Debit")) = "œ«∆‰"
           End If
                    .TextMatrix(Num, .ColIndex("Double_Entry_Vouchers_Description")) = IIf(IsNull(rs("DEV_DES").value), "", Trim(rs("DEV_DES").value))
                Else
                 If val(rs("Credit_Or_Debit").value) = 0 Then
            .TextMatrix(Num, .ColIndex("Credit_Or_Debit")) = "Debit"
           Else
            .TextMatrix(Num, .ColIndex("Credit_Or_Debit")) = "Credit"
           End If
           
                    .TextMatrix(Num, .ColIndex("Double_Entry_Vouchers_Description")) = IIf(IsNull(rs("DEV_DES").value), "", Trim(rs("DEV_DES").value))
                End If
           

                .TextMatrix(Num, .ColIndex("DEV_Value")) = IIf(IsNull(rs("DEV_Value").value), "", Trim(rs("DEV_Value").value))
                  .TextMatrix(Num, .ColIndex("NoteDate")) = IIf(IsNull(rs("NoteDate").value), "", Trim(rs("NoteDate").value))
          
                ' .TextMatrix(Num, .ColIndex("countryid")) = IIf(IsNull(rs("countryid").value), "", (rs("countryid").value))
                '    .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", Trim(rs("CountryName").value))
            
            End With

            rs.MoveNext
        Next Num

        ' Fg.AutoSize 0, Fg.Cols - 1, False
    End If
Else
  MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
End If
    Exit Sub
ErrTrap:
End Sub
Private Sub ALLButton4_Click()
 
    Dim StrAccountCode As String
    Dim StrAccountName As String
    Dim des As String
    
    Dim cAccountReport As New ClsAccReports
    cAccountReport.BegineDate = Me.DTP_Date.value
    cAccountReport.EndDate = Me.DTP_Date1.value

    If Text1(2).Text <> "" Then
        StrAccountCode = Get_Account_code(Text1(2).Text)
        StrAccountName = Get_Account_name(Text1(2).Text)
        des = Text1(1).Text
    Else

        If Adodc1.Recordset.RecordCount > 0 Then
            StrAccountCode = Get_Account_code(Adodc1.Recordset.Fields!account_serial)
            StrAccountName = Get_Account_name(Adodc1.Recordset.Fields!account_serial)
            des = Text1(1).Text
                
        End If
       
    End If
        
    cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, des
        
    Set cAccountReport = Nothing
End Sub



Private Sub ALLButton6_Click()
  If case_id.Caption = 3 Then
  buildSqlNote
  Else
  buildSql
  End If
End Sub
Function buildSqlNote()

sql = "SELECT     TOP 100 PERCENT dbo.Notes1.foxy_no, dbo.Notes1.KALEB, dbo.Notes1.DAWRY, dbo.Notes1.NoteID, dbo.Notes1.NoteType, dbo.Notes1.NoteDate, "
sql = sql & "                      dbo.Notes1.Note_Value  , dbo.Notes1.NoteHijriDate, dbo.Notes1.Remark, dbo.Notes1.general_cost_center, dbo.Notes1.NotePosted, dbo.Notes1.UserID,"
sql = sql & "                      dbo.Notes1.NoteSerial, dbo.Notes1.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_ID,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.UserID AS Expr1, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS1.[Value] as DEV_Value,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS1.valuee, dbo.DOUBLE_ENTREY_VOUCHERS1.currency,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.rate, dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description as DEV_DES,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione, dbo.ACCOUNTS.Account_Name,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial, dbo.Notes1.branch_no,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1.project_id , dbo.Projects.Project_name, dbo.Projects.Project_nameE , dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID"
sql = sql & " FROM         dbo.ACCOUNTS INNER JOIN"
sql = sql & "                      dbo.Notes1 INNER JOIN"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS1 ON dbo.Notes1.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID ON"
sql = sql & "                      dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS1.project_id = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id"
sql = sql & " Where (dbo.Notes1.NoteType = 101) And (Not (dbo.DOUBLE_ENTREY_VOUCHERS1.notes_id Is Null))"
                      
                      

If (Text1(0).Text) <> "" Then
'sql = sql & " and    dbo.Notes.noteSerial   ='" & (Text1(0).text) & "'"
sql = sql & " and ( CONVERT(varchar(100), CAST(dbo.Notes1.noteSerial AS decimal(38, 0))) LIKE '%" & Text1(0).Text & "%')"
'sql = sql & " and (convert(varchar(100),cast( dbo.Notes.noteSerial as decimal(38.0)))) like '% " & Text1(0).text & " % '"

End If


If (Text1(1).Text) <> "" Then
sql = sql & " and    Double_Entry_Vouchers_Description like '%" & Text1(1).Text & "%'"

End If

If (Text1(2).Text) <> "" Then
sql = sql & " and    account_serial like '%" & Text1(2).Text & "%'"

End If

 
If val((Text1(3).Text)) <> 0 Then
sql = sql & " AND value >=" & val(Text1(3).Text)

End If
 
If val((Text1(7).Text)) <> 0 Then
sql = sql & " AND  value <=" & val(Text1(7).Text)

End If

 If (Text1(4).Text) <> "" Then
sql = sql & " and    Notes1.remark like '%" & Text1(4).Text & "%'"

End If
 If (Text1(5).Text) <> "" Then
'sql = sql & " and    Notes.remark like '%" & Text1(5).text & "%'"
sql = sql & " and ( CONVERT(varchar(100), CAST(dbo.Notes1.NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(5).Text & "%')"

End If
' If (Text1(6).text) <> "" Then
'sql = sql & " and    Notes.ManualNo like '%" & Text1(6).text & "%'"
'sql = sql & " and ( CONVERT(varchar(100), CAST(dbo.Notes.NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(6).text & "%')"

'End If

If Check3.value = vbChecked Then
    
    
        sql = sql + " and RecordDate >=" & SQLDate(DTP_Date, True) & ""
        
 
    
End If

If Check4.value = vbChecked Then
    
    
        sql = sql & " and RecordDate <=" & SQLDate(DTP_Date1, True) & ""
        
    
End If


If Check5.value = vbChecked Then
    
    
        sql = sql & " and DueDate >=" & SQLDate(DtpChequeDueDatefrom, True) & ""
        
    
End If

If Check7.value = vbChecked Then
    
    
        sql = sql & " and DueDate <=" & SQLDate(DtpChequeDueDateTo, True) & ""
        
    
End If


 

 
 If Check6.value = vbChecked Then
    
    
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
        
    
End If

 If (TxtChequeNumber.Text) <> "" Then
sql = sql & " and    ChqueNum like '%" & TxtChequeNumber.Text & "%'"

End If

 If (txtperson.Text) <> "" Then
sql = sql & " and    Notes1.person like '%" & txtperson.Text & "%'"

End If



 If DcboBankName.BoundText <> "" Then
    
    
        sql = sql & " and Notes1.BankID =" & val(DcboBankName.BoundText) & ""
        
    
End If



       Retrive sql
   
 

End Function
Function buildSql()


sql = "select * from RptLedger_Sub where NoteType=53 "
sql = "SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, "
sql = sql & "  dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
sql = sql & "  dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
sql = sql & "  dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
sql = sql & "                        dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.Posted,"
sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
sql = sql & "                        dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
sql = sql & "                        dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch,"
sql = sql & "                        dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center, dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters,"
sql = sql & "                        dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS.project_id, dbo.TblNotesTypes.NotesTypeNamee,"
sql = sql & "                        dbo.TransactionTypes.TransactionEnglishName, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.Notes.NoteSerial1,"
sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.Notes.BankID,"
sql = sql & "                        dbo.Notes.person , dbo.Notes.dueDate"
sql = sql & "  FROM         dbo.TblBranchesData INNER JOIN"
sql = sql & "                        dbo.TblUsers INNER JOIN"
sql = sql & "                        dbo.ACCOUNTS INNER JOIN"
sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code ON"
sql = sql & "                        dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
sql = sql & "                        dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id LEFT OUTER JOIN"
sql = sql & "                        dbo.Notes LEFT OUTER JOIN"
sql = sql & "                        dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
sql = sql & "                        dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                        dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type where NoteType=53"
                      
                      

If (Text1(0).Text) <> "" Then
'sql = sql & " and    dbo.Notes.noteSerial   ='" & (Text1(0).text) & "'"
sql = sql & " and ( CONVERT(varchar(100), CAST(dbo.Notes.noteSerial AS decimal(38, 0))) LIKE '%" & Text1(0).Text & "%')"
'sql = sql & " and (convert(varchar(100),cast( dbo.Notes.noteSerial as decimal(38.0)))) like '% " & Text1(0).text & " % '"

End If


If (Text1(1).Text) <> "" Then
sql = sql & " and    Double_Entry_Vouchers_Description like '%" & Text1(1).Text & "%'"

End If

If (Text1(2).Text) <> "" Then
sql = sql & " and    account_serial like '%" & Text1(2).Text & "%'"

End If

 
If val((Text1(3).Text)) <> 0 Then
sql = sql & " AND value >=" & val(Text1(3).Text)

End If
 
If val((Text1(7).Text)) <> 0 Then
sql = sql & " AND  value <=" & val(Text1(7).Text)

End If

 If (Text1(4).Text) <> "" Then
sql = sql & " and    Notes.remark like '%" & Text1(4).Text & "%'"

End If
 If (Text1(5).Text) <> "" Then
'sql = sql & " and    Notes.remark like '%" & Text1(5).text & "%'"
sql = sql & " and ( CONVERT(varchar(100), CAST(dbo.Notes.NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(5).Text & "%')"

End If
 If (Text1(6).Text) <> "" Then
sql = sql & " and    Notes.ManualNo like '%" & Text1(6).Text & "%'"
'sql = sql & " and ( CONVERT(varchar(100), CAST(dbo.Notes.NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(6).text & "%')"

End If

If Check3.value = vbChecked Then
    
    
        sql = sql + " and RecordDate >=" & SQLDate(DTP_Date, True) & ""
        
 
    
End If

If Check4.value = vbChecked Then
    
    
        sql = sql & " and RecordDate <=" & SQLDate(DTP_Date1, True) & ""
        
    
End If


If Check5.value = vbChecked Then
    
    
        sql = sql & " and DueDate >=" & SQLDate(DtpChequeDueDatefrom, True) & ""
        
    
End If

If Check7.value = vbChecked Then
    
    
        sql = sql & " and DueDate <=" & SQLDate(DtpChequeDueDateTo, True) & ""
        
    
End If


 

 
 If Check6.value = vbChecked Then
    
    
        sql = sql & " and branch_id =" & val(dcBranch.BoundText) & ""
        
    
End If

 If (TxtChequeNumber.Text) <> "" Then
sql = sql & " and    ChqueNum like '%" & TxtChequeNumber.Text & "%'"

End If

 If (txtperson.Text) <> "" Then
sql = sql & " and    notes.person like '%" & txtperson.Text & "%'"

End If



 If DcboBankName.BoundText <> "" Then
    
    
        sql = sql & " and notes.BankID =" & val(DcboBankName.BoundText) & ""
        
    
End If



       Retrive sql
   
 

End Function

Private Sub ALLButton7_Click()
clear_all Me
End Sub

Private Sub Combo1_Click()

  '  If Combo1.ListIndex = 0 Then
  '      sql = "SELECT     * from dbo.RptLedger_Sub  where NoteType=53 "
  '
  '  ElseIf Combo1.ListIndex = 1 Then
  '      sql = "SELECT     * from dbo.RptLedger_Sub  where NoteType<>53 "
  '
  '  End If

End Sub

Private Sub Command1_Click()

    If replace_in_data_base("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_Description", txt1.Text, txt2.Text) = True Then
        MsgBox "Done"
      
        ALLButton2_Click
    End If

End Sub

Private Sub DataGrid2_Click()
  
With DataGrid2
If case_id.Caption = 2 Then
FrmAccEditJournal3.TxtModFlg = "R"
        FrmAccEditJournal3.Retrive .TextMatrix(.Row, .ColIndex("NoteSerial"))
        FrmAccEditJournal3.show
     ElseIf case_id.Caption = 3 Then
     FrmAccEditJournal1.TxtModFlg = "R"

     FrmAccEditJournal1.Retrive val(.TextMatrix(.Row, .ColIndex("Notes_ID")))
        FrmAccEditJournal1.show
     
        'UPDATE_RECORDS
    End If
    End With
End Sub

Private Sub DCExtraAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 270815
            
    End If
    
      If KeyCode = vbKeyReturn Then
      ALLButton6_Click
            
    End If
    
End Sub

Private Sub Form_Activate()

    If first_run = True Then

        Exit Sub
    Else
        first_run = True
 
        sql = "SELECT     * from dbo.RptLedger_Sub   where Double_Entry_Vouchers_ID=0"
      
 
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
       If case_id.Caption = 2 Then
     LblHeader.Caption = "«·»ÕÀ ⁄‰ ”‰œ ’—› „ ⁄œœ"
      ElseIf case_id.Caption = 3 Then
     LblHeader.Caption = "«·»ÕÀ ⁄‰ «·ﬁÌÊœ «·«›  «ÕÌ…"
    End If
    Else
        If case_id.Caption = 2 Then
     LblHeader.Caption = "Search Multiple Payments Voucher"
      ElseIf case_id.Caption = 3 Then
     LblHeader.Caption = "Search Opening Balance"
    End If
    End If
    Me.Caption = LblHeader.Caption


End Sub

Private Sub ChangeLang()
    Me.Caption = "Voucher Search"
    ALLButton5.Caption = "Save"
    Label3.Caption = "ACC Code"
    OpTcode(0).Caption = "Acc"
        OpTcode(1).Caption = "Customer"
            OpTcode(2).Caption = "Supp."
                OpTcode(3).Caption = "Employee"
              Label13.Caption = "To Value"
              Label11.Caption = "Vchr No"
              ALLButton7.Caption = "Clear"
                    ALLButton6.Caption = "Search"
    Frame3.Caption = "Update Description Per Line "
    lbl(16).Caption = "Bank Name"
    lbl(34).Caption = "Beneficiary Name"
    lbl(17).Caption = "Check No"
    lbl(18).Caption = "Date"
    
    Check6.RightToLeft = False
    Check6.Caption = "Branch"
    LblHeader.Caption = Me.Caption
    Label4.Caption = "Search"
    Label6.Caption = "Voucher#"
    Label7.Caption = "General Des"
    Label1.Caption = "Des"
    Label2.Caption = "From Value"
    Check7.RightToLeft = False
    Check7.Caption = "To"
    Check5.RightToLeft = False
    Check5.Caption = "From"
        With DataGrid2
    .TextMatrix(0, .ColIndex("NoteSerial1")) = "Vchr No"
    .TextMatrix(0, .ColIndex("ManualNo")) = "Manual No"
    .TextMatrix(0, .ColIndex("NoteDate")) = "Date"
    .TextMatrix(0, .ColIndex("NoteSerial")) = "GL No"
    .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Code"
    .TextMatrix(0, .ColIndex("DEV_Value")) = "Value"
    .TextMatrix(0, .ColIndex("Credit_Or_Debit")) = "Credit_Or_Debit"
    .TextMatrix(0, .ColIndex("Double_Entry_Vouchers_Description")) = "Description"
    End With
    Label5.Caption = "Source"
    Combo1.Clear
    Combo1.AddItem "Manual"
    Combo1.AddItem "Auto"
    Label9.Caption = "Account#"
    Frame1.Caption = "Date"
    Check3.Caption = "From"
    Label12.Caption = "Manual No"
    Check4.Caption = "To"
    ALLButton2.Caption = "Search"
    ALLButton4.Caption = "Print Statement of Acc."
    Frame2.Caption = "Replace Text "
    Label8.Caption = "Find"
    Label10.Caption = "Replace By"
    Command1.Caption = "Replace All"
    DataGrid2.Columns(3).Caption = "Voucher#"
    DataGrid2.Columns(8).Caption = "Account#"
    DataGrid2.Columns(9).Caption = "Value"
    DataGrid2.Columns(11).Caption = "Des"
    DataGrid2.Columns(12).Caption = "date"
    ALLButton1.Caption = "Ok"


End Sub

Private Sub Form_Load()
    On Error Resume Next
    DTP_Date.value = Date
    DTP_Date1.value = Date
         
    Combo1.Clear
    Combo1.AddItem "ÌœÊÌ"
    Combo1.AddItem "«·Ì"
 
     
    '

    Me.left = (mdifrmmain.width - Me.width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
 
       Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
      Dcombos.GetAccountingCodes DCExtraAccount, True
       Dcombos.GetBranches dcBranch
       Dcombos.GetBanks DcboBankName
       
       
      
    If my_language = "E" Then
        DataGrid1.Visible = True
        DataGrid2.Visible = False
    End If
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    first_run = False
End Sub

Private Sub Text1_Change(Index As Integer)
     DCExtraAccount.BoundText = ""
       DCExtraAccount.BoundText = Get_Account_code(Text1(2).Text, 1)

End Sub

Private Sub Text1_KeyUp(Index As Integer, _
                        KeyCode As Integer, _
                        Shift As Integer)

    'On Error Resume Next
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 3

    End If

    If KeyCode = 13 Then
ALLButton6_Click
        '     On Error Resume Next
             
'        If Index = 0 Then
'
'            If Check1.value = 1 Then
'
'                sql = "select * from RptLedger_Sub where    noteSerial ='" & val(Text1(Index).text) & "'"
'            Else
'               ' sql = "select * from RptLedger_Sub where    noteSerial ='" & val(Text1(Index).text) & "'"
'                sql = "select * from RptLedger_Sub where ( CONVERT(varchar(100), CAST(noteSerial AS decimal(38, 0))) LIKE '%" & Text1(Index).text & "%')"
'
                'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
'            End If
'        End If
'           If Index = 6 Then
'
'            If Check1.value = 1 Then
'
'                sql = "select * from RptLedger_Sub where    ManualNo ='" & (Text1(Index).text) & "'"
'            Else
'                sql = "select * from RptLedger_Sub where    ManualNo ='" & (Text1(Index).text) & "'"
'
  '              'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
  '          End If
  '      End If
 ''
  '           If Index = 5 Then
 '
  '          If Check1.value = 1 Then
                        
'                sql = "select * from RptLedger_Sub where    NoteSerial1 ='" & Text1(Index).text & "'"
'            Else
'               ' sql = "select * from RptLedger_Sub where    NoteSerial1 ='" & Text1(Index).text & "'"
'                sql = "select * from RptLedger_Sub where ( CONVERT(varchar(100), CAST(NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(Index).text & "%')"
'
    '            'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
    '        End If
    '    End If
    '
    '    If Index = 1 Then
    '
    '        sql = "select * from RptLedger_Sub where    DEV_DES like '%" & Text1(Index).text & "%'"
    '        '    If Check1.value = 1 Then
    '        '      Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_type = '" & Text1(Index).text & "' "
    '        '     Else
    '        '      Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_type like '%" & Text1(Index).text & "%'"
    '        '
   ''
   '         '                  End If
   '     End If
   '
   '     If Index = 2 Then
   '         sql = "select * from RptLedger_Sub where account_serial like '%" & Text1(2).text & "%' and   DEV_DES like '%" & Text1(1).text & "%'"
   '
   '     End If
   '
   '     If Index = 3 Then
   '         If Not IsNumeric(Text1(Index).text) Then Exit Sub
   '         sql = "select * from RptLedger_Sub where dev_value =" & val(Text1(Index).text)
   '     End If
   '
   '     If Index = 4 Then
   '         'If Check1.value = 1 Then
   '         sql = "select * from RptLedger_Sub where    remark like '%" & Text1(Index).text & "%'"
   '         'Else
   '         ' Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and description like '%" & Text1(Index).text & "%'"
   '
            'End If
   '     End If
                  
       ' If Index = 5 Then
       '     If Check1.value = 1 Then
                '            Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_source = '" & Text1(Index).text & "' "
       '     Else
                '            Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_source like '%" & Text1(Index).text & "%'"
    
       '     End If
       ' End If
             
   ' Retrive sql

   ' buildSql

       
 
    End If

End Sub

