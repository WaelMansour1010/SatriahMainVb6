VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Voucher_search2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ”‰œ  ﬁÌÊœ «· ”ÊÌ… «·ÌœÊÌ…"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13830
   Icon            =   "Voucher_search2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   13830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame5 
      Height          =   4215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   3480
      Width           =   13815
      Begin VSFlex8UCtl.VSFlexGrid DataGrid2 
         Height          =   3645
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   13575
         _cx             =   23945
         _cy             =   6429
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
         FormatString    =   $"Voucher_search2.frx":000C
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
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1560
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
         MICON           =   "Voucher_search2.frx":0196
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
         MICON           =   "Voucher_search2.frx":01B2
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
      Height          =   975
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2520
      Width           =   3495
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ï"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTP_Date 
         Height          =   330
         Left            =   840
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
         Format          =   94240771
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Date1 
         Height          =   330
         Left            =   840
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94240771
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
         MICON           =   "Voucher_search2.frx":01CE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
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
      MICON           =   "Voucher_search2.frx":01EA
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
      Left            =   2400
      TabIndex        =   19
      Tag             =   "Press F3 To Search"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   4
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   " Õ ÊÏ «·ﬂ·„…"
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "«·ﬂ·„… ›ﬁÿ"
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2760
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
      Visible         =   0   'False
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
      MICON           =   "Voucher_search2.frx":0206
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
      Top             =   480
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
      MICON           =   "Voucher_search2.frx":0222
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton6 
      Height          =   615
      Left            =   120
      TabIndex        =   49
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
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
      MICON           =   "Voucher_search2.frx":023E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton7 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   50
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "„”Õ"
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
      MICON           =   "Voucher_search2.frx":025A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Height          =   255
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„’œ—"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„»·€"
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
      Caption         =   "«·»ÕÀ ⁄‰ ”‰œ  ﬁÌÊœ «· ”ÊÌ… «·ÌœÊÌ…"
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
      Left            =   12480
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ «·⁄«„"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
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
      Visible         =   0   'False
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
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Voucher_search2"
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
     sql = sql & " and  NoteType=57 "
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
               .TextMatrix(Num, .ColIndex("ManualNo")) = IIf(IsNull(rs("ManualNo").value), "", Trim(rs("ManualNo").value))

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
GetData
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  Dim RptLedger_Sub As String
     
     RptLedger_Sub = "SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, "
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.NoteDate, dbo.Notes.NoteType,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Parent_Account_Code,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch, dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters, dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblNotesTypes.NotesTypeNamee, dbo.TransactionTypes.TransactionEnglishName, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Posted, dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.currency,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.TblBranchesData.RegionID, dbo.TblSection.name, dbo.TblSection.namee,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount, dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code, dbo.DOUBLE_ENTREY_VOUCHERS.project_id,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.DOUBLE_ENTREY_VOUCHERS.operid,"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.pandid , dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo"
RptLedger_Sub = RptLedger_Sub & " FROM         dbo.TblAqar RIGHT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblBranchesData INNER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblUsers INNER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id ON"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid LEFT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.Notes LEFT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
RptLedger_Sub = RptLedger_Sub & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
 

        
        
    'sql = "select * from RptLedger_Sub"
       sql = RptLedger_Sub
       
       BolBegine = False
       StrWhere = " where 1=1"
BolBegine = True
  If Me.Text1(5).Text <> "" Then
        If BolBegine = True Then
            'StrWhere = StrWhere & "AND   NoteSerial1 = " & val(Text1(5).text) & ""
            StrWhere = StrWhere & "  AND ( CONVERT(varchar(100), CAST(Notes.NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(5).Text & "%')"
        Else
            BolBegine = True
           ' StrWhere = " Where  NoteSerial1 = " & val(Text1(5).text) & ""
            StrWhere = "  where ( CONVERT(varchar(100), CAST(Notes.NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(5).Text & "%')"
        End If
    End If
      If Me.Text1(6).Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND   Notes.ManualNo = '" & Text1(6).Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  Notes.ManualNo = '" & Text1(6).Text & "'"
        End If
    End If
          If Me.Text1(0).Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND   Notes.noteSerial = '" & Text1(0).Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  Notes.noteSerial = '" & Text1(0).Text & "'"
        End If
    End If
              If val(Me.Text1(3).Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND   value = " & val(Text1(3).Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where  value =  " & val(Text1(3).Text) & ""
        End If
    End If
             If Me.Text1(4).Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND   Notes.remark like '%" & Text1(4).Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  Notes.remark like '%" & Text1(4).Text & "%'"
        End If
    End If
             If Me.Text1(1).Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND   Double_Entry_Vouchers_Description like '%" & Text1(1).Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  Double_Entry_Vouchers_Description like '%" & Text1(1).Text & "%'"
        End If
    End If
                 If Me.Text1(2).Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND   account_serial like '%" & Text1(2).Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  account_serial like '%" & Text1(2).Text & "%'"
        End If
    End If
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Check3.value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND recorddate >=" & SQLDate(Me.DTP_Date.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where recorddate>=" & SQLDate(Me.DTP_Date.value, True) & ""
        End If
    End If
    If Check4.value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND recorddate <=" & SQLDate(Me.DTP_Date1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  recorddate <=" & SQLDate(Me.DTP_Date1.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   
    sql = sql & StrWhere
Retrive sql
 End Sub

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
If case_id.Caption = 1 Then
FrmAccEditJournal4.TxtModFlg = "R"
        FrmAccEditJournal4.Retrive .TextMatrix(.Row, .ColIndex("Notes_ID"))
        FrmAccEditJournal4.show
      
        'UPDATE_RECORDS
    End If
    End With
End Sub

Private Sub Form_Activate()

    If first_run = True Then

        Exit Sub
    Else
        first_run = True
 
        sql = "SELECT     * from dbo.RptLedger_Sub   where Double_Entry_Vouchers_ID=0"
      
 
    End If

End Sub

Private Sub ChangeLang()
    Me.Caption = "Voucher Search"
    ALLButton5.Caption = "Save"
    Label3.Caption = "ACC Code"
    OpTcode(0).Caption = "Acc"
        OpTcode(1).Caption = "Customer"
            OpTcode(2).Caption = "Supp."
                OpTcode(3).Caption = "Employee"
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
  ALLButton7.Caption = "Clear"
  Label11.Caption = "Vchr No"
  ALLButton6.Caption = "Search"
  Check1.RightToLeft = False
  Check1.Caption = "Only Word"
    Label12.Caption = "Manual No"
    Frame3.Caption = "Update Description Per Line "
    LblHeader.Caption = Me.Caption
    Label4.Caption = "Search"
    Label6.Caption = "Voucher#"
    Label7.Caption = "General Des"
    Label1.Caption = "Des"
    Label2.Caption = "Value"
    Label5.Caption = "Source"
    Combo1.Clear
    Combo1.AddItem "Manual"
    Combo1.AddItem "Auto"
    Label9.Caption = "Account#"
    Frame1.Caption = "Date"
    Check3.Caption = "From"
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

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
 
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

Private Sub Text1_KeyUp(Index As Integer, _
                        KeyCode As Integer, _
                        Shift As Integer)

    'On Error Resume Next
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 3

    End If

    If KeyCode = 13 Then
GetData
End If
        '     On Error Resume Next
     '
     '   If Index = 0 Then
 '
 '           If Check1.value = 1 Then
 '
 '               sql = "select * from RptLedger_Sub where    noteSerial ='" & val(Text1(Index).text) & "'"
 '           Else
'                sql = "select * from RptLedger_Sub where    noteSerial ='" & val(Text1(Index).text) & "'"
'
'                'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
'            End If
'        End If
'           If Index = 6 Then
'
'            If Check1.value = 1 Then
'
'                sql = "select * from RptLedger_Sub where    ManualNo ='" & (Text1(Index).text) & "'"
'            Else
 '               sql = "select * from RptLedger_Sub where    ManualNo ='" & (Text1(Index).text) & "'"
'
'                'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
 '           End If
''        End If
        
 '            If Index = 5 Then
 '
 '           If Check1.value = 1 Then
 '
 '               sql = "select * from RptLedger_Sub where    NoteSerial1 = '" & Text1(Index).text & "'"
 '           Else
 '               sql = "select * from RptLedger_Sub where ( CONVERT(varchar(100), CAST(NoteSerial1 AS decimal(38, 0))) LIKE '%" & Text1(Index).text & "%')"
 '

 '               'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
 '           End If
 '       End If
 '
 '       If Index = 1 Then
 '
 '           sql = "select * from RptLedger_Sub where    DEV_DES like '%" & Text1(Index).text & "%'"
 '           '    If Check1.value = 1 Then
 '           '      Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_type = '" & Text1(Index).text & "' "
 '           '     Else
 '           '      Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_type like '%" & Text1(Index).text & "%'"
 '           '
 '
 '           '                  End If
 '       End If
 '
 '       If Index = 2 Then
 '           sql = "select * from RptLedger_Sub where account_serial like '%" & Text1(2).text & "%' and   DEV_DES like '%" & Text1(1).text & "%'"
 '
 '       End If
 '
 '       If Index = 3 Then
 '           If Not IsNumeric(Text1(Index).text) Then Exit Sub
 '           sql = "select * from RptLedger_Sub where dev_value =" & val(Text1(Index).text)
 '       End If
 '
 '       If Index = 4 Then
 '           'If Check1.value = 1 Then
 '           sql = "select * from RptLedger_Sub where    remark like '%" & Text1(Index).text & "%'"
 '           'Else
 '           ' Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and description like '%" & Text1(Index).text & "%'"
 '
 '           'End If
 '       End If
 '
       ' If Index = 5 Then
 '      '     If Check1.value = 1 Then
 '               '            Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_source = '" & Text1(Index).text & "' "
 '      '     Else
 '               '            Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_source like '%" & Text1(Index).text & "%'"
 '
 '      '     End If
 '      ' End If
 '
 '   Retrive sql
'
      

       
 
'    End If

End Sub

