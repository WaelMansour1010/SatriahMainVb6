VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form CashierLogout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈€·«ﬁ ««·‘Ì›  "
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9825
   Icon            =   "CashierLogout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   9825
   Begin VB.Frame Frame6 
      Height          =   2175
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   5640
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Frame Frame4 
      Caption         =   " Õ·Ì·Ï «·„»Ì⁄« "
      Height          =   6015
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   600
      Width           =   3255
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1440
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92930049
         CurrentDate     =   38784
      End
      Begin VSFlex8UCtl.VSFlexGrid Grid 
         Height          =   3765
         Left            =   0
         TabIndex        =   48
         Top             =   600
         Width           =   3045
         _cx             =   5371
         _cy             =   6641
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
         BackColor       =   -2147483640
         ForeColor       =   65280
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483641
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483640
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"CashierLogout.frx":000C
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
         Editable        =   0
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
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92930049
         CurrentDate     =   38784
      End
      Begin VB.Label LblNet 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " ’«›Ì «·‰ﬁÿ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   1560
         TabIndex        =   55
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì „—œÊœ«  «·‰ﬁÿ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   1560
         TabIndex        =   54
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label lblReturn 
         Alignment       =   2  'Center
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lblTotalTransaction 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì „»Ì⁄«  «·‰ﬁÿ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   49
         Top             =   4440
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "»Ì«‰«  „Õ«”»Ì… "
      Height          =   4095
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame5 
         Height          =   2175
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.TextBox Txtrecivedpetty 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtPosSales 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox XPTxtPass 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2520
         Width           =   3225
      End
      Begin VB.TextBox TxtBalance 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   360
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ALLButtonS.ALLButton CMDLogin 
         Height          =   375
         Left            =   2040
         TabIndex        =   30
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " √ﬂÌœ «·œŒÊ·"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogout.frx":01F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton CMDCancel 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«·€«¡"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   192
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogout.frx":0210
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   495
         Left            =   4080
         TabIndex        =   58
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16384
         BCOLO           =   16384
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogout.frx":022C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton4 
         Height          =   495
         Left            =   2040
         TabIndex        =   60
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   " ﬁ—Ì— «·ÌÊ„Ì…"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16384
         BCOLO           =   16384
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogout.frx":0248
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton5 
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   3480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   " ﬁ—Ì—«·ÌÊ„"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16384
         BCOLO           =   16384
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogout.frx":0264
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblTotalsalecash 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LBLtTALCASHES 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " ’«›Ì «·‰ﬁœÌ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   4200
         TabIndex        =   51
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label LblTotalCollected 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label LblTotals 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì  "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   1440
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì «·„ÿ·Ê»"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   4200
         TabIndex        =   42
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„” ·„"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„” ·„"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   39
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄Âœ… ”«»ﬁ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label LblPetty 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LBLBalance 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "—’Ìœ ”«»ﬁ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„” ·„"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·ﬂ«‘Ì—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   25
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ·„… «·„—Ê—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   24
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   3360
         Picture         =   "CashierLogout.frx":0280
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«”„ «·‰ﬁÿ…"
      Height          =   615
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   6135
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcpoint 
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcShift 
         Height          =   360
         Left            =   5760
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õœœ «·‘Ì› "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "»Ì«‰«  «·‰ﬁÿ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  «·„‘—›"
      Height          =   1455
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   6135
      Begin VB.TextBox XPTxtPass1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   3225
      End
      Begin MSDataListLib.DataCombo DCboUserName1 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " √ﬂÌœ «·œŒÊ·"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "CashierLogout.frx":095D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„‘—›"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ·„… «·„—Ê—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   3360
         Picture         =   "CashierLogout.frx":0979
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
   End
   Begin MSComCtl2.DTPicker ShfitFrom 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "'Time: 'hh:mm tt"
      Format          =   92930051
      UpDown          =   -1  'True
      CurrentDate     =   39240
   End
   Begin MSComCtl2.DTPicker ShfitTo 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "'Time: 'hh:mm tt"
      Format          =   92930051
      UpDown          =   -1  'True
      CurrentDate     =   39240
   End
   Begin MSDataListLib.DataCombo DcboDebitSide 
      Height          =   360
      Left            =   6120
      TabIndex        =   32
      Top             =   7560
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16761024
      ForeColor       =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DcboCreditSide 
      Height          =   360
      Left            =   3840
      TabIndex        =   33
      Top             =   8280
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16761024
      ForeColor       =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   615
      Left            =   120
      TabIndex        =   46
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "«” ·«„ «·„»«·€  Ê› Õ «·‘Ì›  «·À«‰Ì"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16384
      BCOLO           =   16384
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "CashierLogout.frx":1056
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   0
      Picture         =   "CashierLogout.frx":1072
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   660
   End
   Begin VB.Label LBLShiftID 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
   Begin VB.Label LBLPOSName 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label LBLPOSCode 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4320
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ì«‰«  «·‘Ì› "
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·‰ﬁÿ…"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "≈€·«ﬁ ««·‘Ì› "
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
      Height          =   585
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9555
   End
   Begin VB.Label Labelx 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·‰ﬁÿ…"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   11520
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "CashierLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PointID As Integer
Dim Pointname As String
Dim Balance As Double
Dim pettyBalance As Double
Dim cashierBocaccount As String
Dim SupervisorBocaccount As String
Dim Account_Code_dynamic  As String
Dim Account_Code_dynamic1  As String
Public ScreenJob As Integer
Private Sub ChangeLang()
Me.Caption = "Shift Logout"
''Shift Logou
ALLButton4.Caption = "Print Reports"
LblHeader.Caption = Me.Caption
Labelx(21).Caption = "Net"
If ScreenJob = 1 Then
Me.Caption = "Cashier  Logout"
LblHeader.Caption = Me.Caption
 End If
ALLButton3.Caption = "Print"
    Frame1.Caption = "Supervisor"
     Frame2.Caption = "Point"
      Frame3.Caption = "Accounts Data"
      
    Labelx(r).Caption = "Shift"
    Labelx(2).Caption = "From"
    Labelx(3).Caption = "To"
 Labelx(11).Caption = " Point"
 
      
          Labelx(8).Caption = " Shift"
    Labelx(10).Caption = " Name"
    Labelx(9).Caption = "Password"
ALLButton1.Caption = "LogIn"

 
    Labelx(6).Caption = " Name"
    Labelx(7).Caption = "Password"
CMDLogin.Caption = "LogIn"
   
        Labelx(4).Caption = "Balance"
              Labelx(12).Caption = "Pettycash Blance"
              
     Labelx(5).Caption = "Pettycash"
     Labelx(14).Caption = "Payed"
          Labelx(15).Caption = "Payed"
           Labelx(17).Caption = "Total"
          Labelx(18).Caption = "Net Sales"
           Labelx(16).Caption = "Totals"
           
         Labelx(13).Caption = "Total Sales"
         Labelx(20).Caption = "Total Refund"
         Frame4.Caption = "Detailed Sales"
         
          
     CMDCancel.Caption = "Cancel"
     ALLButton2.Caption = "Recive Money and Close the Point"
       
     

  With Grid
'        .TextMatrix(0, .ColIndex("ID")) = "Cashier ID"
        .TextMatrix(0, .ColIndex("PaymentName")) = "Payment"
        .TextMatrix(0, .ColIndex("value")) = "value"
         
 
    End With
      
End Sub
Private Sub ALLButton1_Click()

     On Error GoTo ErrTrap
    If DCboUserName1.Text = "" Then
         If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· «”„ «·„‘—›"
 Else
 Msg = "Select dmin User Name"
 
 
 End If
 
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If
 
 
       If dcpoint.BoundText = "" Then
                   If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ»  ÕœÌœ  «·‰ﬁÿ…"
        Else
        Msg = " Select POS First"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        dcpoint.SetFocus
        Exit Sub
    End If
    

    StrSQL = "Select * From cachierData Where id=" & Me.DCboUserName1.BoundText & " AND password='" & Trim(Me.XPTxtPass1.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
       
       
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " √ﬂœ „‰ ’Õ… «”„ «·„‘—› " & CHR(13)
                Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
            Else
            
            Msg = "User Name Or Password Incorrect " & CHR(13)
            End If


        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName1.SetFocus
        Exit Sub
    End If

    ' user_name = rs("UserName").value
    ' user_id = rs("UserID").value
    ' User_Password = rs("PassWord").value
 
    AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «·œŒÊ· ·‰›ÿ… «·»Ì⁄ »√”„ «·„‘—›  " & DCboUserName1.Text, " System Login", Me.Name, "L", "", ""
  '  AddSessonData val(LBLPOSCode), val(LBLShiftID.Caption), val(DCboUserName.BoundText), ShfitFrom.value, ShfitTo.value, val(LBLBalance.Caption), val(TxtBalance.text), Now
   Frame2.Visible = True
 Frame3.Visible = True
 
'  LBLtTALCASHES.Caption = GetPointSAles(val(dcpoint.BoundText), 0)
LBLtTALCASHES.Caption = GetPointSAles(val(dcpoint.BoundText), 0, 0)
  
lblTotalTransaction.Caption = GetPointSAles(val(dcpoint.BoundText), , 1)
  lblReturn.Caption = GetPointSAles(val(dcpoint.BoundText), , -1)
  
  LblNet.Caption = val(lblTotalTransaction.Caption) + val(lblReturn.Caption)
  
 FillGridWithData val(dcpoint.BoundText)
 
Calc

    Exit Sub
ErrTrap:


End Sub

Private Sub ALLButton2_Click()
     If CheckAcconts = False Then
   Exit Sub
   End If
   
   
              If val(lblTotalTransaction.Caption) > 0 Or val(lblReturn.Caption) > 0 Then
                                                                If CheckWORKINposvATsCREEN = False Then       '«· √ﬂœ „‰ ⁄œ„ «·⁄„· ⁄·Ì ‘«‘Â ﬁÌÊœ «·ﬁÌ„Â «·„÷«›…
                                                                         createVoucher
                                                                End If
    
                    Cn.Execute "   update TblTransactionPayments set locked=1 , lokeddate=" & SQLDate(Now, True) & " where locked is null and PointID=" & PPointID
               End If
    
    
    
    AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «ﬁ›«· ·‰›ÿ… «·»Ì⁄ »√”„  " & DCboUserName.Text, " System Out", Me.Name, "L", "", ""
    AddSessonData val(LBLPOSCode), val(LBLShiftID.Caption), val(DCboUserName.BoundText), ShfitFrom.value, ShfitTo.value, val(LBLBalance.Caption), val(TxtBalance.Text), , Now
'Unload Me
If ScreenJob = 0 Then
'CashierLogin.show
 

End If
Unload Me
'End

End Sub

Private Sub ALLButton3_Click()
print_report
End Sub

Private Sub ALLButton4_Click()
print_ReportDay
End Sub

Private Sub ALLButton5_Click()
'print_reportAll
GetDataNetwork
End Sub
Public Sub GetDataNetwork()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    Dim EmpID As Integer
    DTPicker1.value = Date
    DTPicker1.value = "2021-11-09"
   GetUserData user_id, , , , , , EmpID
    StrSQL = "select * from (( SELECT     dbo.TblTransactionPayments.id, dbo.TblTransactionPayments.Transaction_ID, dbo.TblTransactionPayments.PaymentID, ISNULL(dbo.TblTransactionPayments.[value], "
    StrSQL = StrSQL & "                   dbo.Transactions.Transaction_NetValue) AS Value, dbo.TblTransactionPayments.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
    StrSQL = StrSQL & "                   dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
    StrSQL = StrSQL & "                   dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                   dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                   dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "        FROM         dbo.TblEmployee INNER JOIN"
    StrSQL = StrSQL & "                   dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID)"
    StrSQL = StrSQL & " Union (SELECT     dbo.TblSalesPayment.ID AS id, dbo.TblSalesPayment.TransID AS Transaction_ID, dbo.TblSalesPayment.PaymentID, ISNULL(dbo.TblSalesPayment.[Value], "
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_NetValue) AS value, dbo.TblSalesPayment.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
    StrSQL = StrSQL & "                  dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
    StrSQL = StrSQL & "                  dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                  dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "         FROM         dbo.TblSalesPayment RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id ON"
    StrSQL = StrSQL & "                  dbo.TblSalesPayment.TransID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblSalesPayment.PaymentID = dbo.TblPaymentType.PaymentID where dbo.TblSalesPayment.PaymentID =0)) as x where ( Transaction_Type = 21)"
   
    
        StrSQL = StrSQL & " AND   Emp_ID = " & EmpID & ""
 

    If Not IsNull(Me.DTPicker1.value) Then
        StrSQL = StrSQL & " AND Transaction_Date >=" & SQLDate(Me.DTPicker1.value, True) & ""
    End If
    If Not IsNull(Me.DTPicker1.value) Then
        StrSQL = StrSQL & " AND Transaction_Date<=" & SQLDate(Me.DTPicker1.value, True) & ""
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
      Else
      Msg = "No Data"
    End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportCasher StrSQL, EmpID
End If
End Sub
Function print_reportCasher(Optional NoteSerial As String, Optional EmpID As Integer)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Debug.Print NoteSerial
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCloseCasherAll.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCloseCasherAll.rpt"
         End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo



    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    End If
   xReport.ParameterFields(4).AddCurrentValue GetTimeTrans(GetMinTransID(EmpID))
   xReport.ParameterFields(5).AddCurrentValue GetTimeTrans(GetMaxTransID(EmpID))
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    
    CViewer.FireReport xReport, PrinterTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CMDLogin_Click()

    'On Error GoTo ErrTrap
    If DCboUserName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· «”„ «·„” Œœ„"
 Else
 Msg = "Select User Name"
 End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If

    If XPTxtPass.Text = "" Then
        '    Msg = "„‰ ›÷·ﬂ √œŒ· ﬂ·„… «·„—Ê—"
        '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    XPTxtPass.SetFocus
        '    Exit Sub
    End If

    StrSQL = "Select * From cachierData Where id=" & Me.DCboUserName.BoundText & " AND password='" & Trim(Me.XPTxtPass.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
         
         
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " √ﬂœ „‰ ’Õ… «”„ «·„” Œœ„ " & CHR(13)
                Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
            Else
            
            Msg = "User Name Or Password Incorrect " & CHR(13)
            End If


        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If

    ' user_name = rs("UserName").value
    ' user_id = rs("UserID").value
    ' User_Password = rs("PassWord").value

 lblTotalTransaction.Caption = GetPointSAles(val(dcpoint.BoundText))
 LBLtTALCASHES.Caption = GetPointSAles(val(dcpoint.BoundText), 0, 0)
 
  
 FillGridWithData val(dcpoint.BoundText)
Calc
    
 ALLButton2.Visible = True
 
    Exit Sub
ErrTrap:

End Sub

Private Sub createVoucher()
Dim des As String
Dim DebitAccount As String
Dim CreditAccount As String
DebitAccount = DcboDebitSide.BoundText
CreditAccount = DcboCreditSide.BoundText
des = "   ”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄   ··ﬂ«‘Ì—  " & DCboUserName1.Text & "   ··‰ﬁÿ… " & dcpoint.Text & "   »√‘—«› " & DCboUserName.Text & " «Ã„«·Ì «·„ÿ·Ê» " & LblTotals.Caption & " «Ã„«·Ì «·„”œœ " & LblTotalCollected.Caption
Dim NoteID As Long
 
 
 CreateNotes NoteID, Date, Current_branch, 64, val(LblTotalCollected.Caption)
         
       CREATE_VOUCHER_GE NoteID, Current_branch, user_id, val(TxtBalance.Text), DebitAccount, CreditAccount, des, Date


End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, Notevalue As Double, DebitAccount As String, CreditAcc As String, des As String, NoteDate As Date)
  
 
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
   Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
Dim PaymentName As String
   
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

 '***********************«À»«  „»Ì⁄«  «·‰ﬁÿ… *****************************
 If val(lblTotalTransaction.Caption) > 0 Then     '
    
 ' sql = "SELECT     SUM(dbo.TblTransactionPayments.[value] *isnull(effect,1) ) AS totals, dbo.TblPaymentType.Accountsus, dbo.TblBoxesData.Account_Code, "
'sql = sql & "   dbo.TblTransactionPayments.PaymentID"
'sql = sql & " , dbo.TblPaymentType.PaymentName  FROM         dbo.TblTransactionPayments INNER JOIN"
'sql = sql & " dbo.cachierData ON dbo.TblTransactionPayments.CurrentCashireID = dbo.cachierData.id INNER JOIN"
'sql = sql & " dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
'sql = sql & " dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
'sql = sql & " Where (dbo.TblTransactionPayments.locked Is Null) And (dbo.TblTransactionPayments.PointID = " & PPointID & ")"


         'sql = sql & " AND  (dbo.TblTransactionPayments.Effect= 1  or dbo.TblTransactionPayments.Effect is null )"
'         sql = sql & " GROUP BY dbo.TblPaymentType.Accountsus, dbo.TblBoxesData.Account_Code, dbo.TblTransactionPayments.PaymentID, dbo.TblPaymentType.PaymentName"
'sql = sql & " ORDER BY dbo.TblTransactionPayments.PaymentID"
 
   
 sql = " SELECT     TOP 100 PERCENT SUM(dbo.TblTransactionPayments.[value] * ISNULL(dbo.TblTransactionPayments.Effect, 1)) AS totals, dbo.TblPaymentType.Accountsus, "
sql = sql & "                       dbo.TblTransactionPayments.PaymentID , dbo.TblPaymentType.PaymentName"
sql = sql & " FROM         dbo.TblTransactionPayments LEFT OUTER JOIN"
sql = sql & "                      dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
sql = sql & " Where (dbo.TblTransactionPayments.locked Is Null) And (dbo.TblTransactionPayments.PointID = " & val(dcpoint.BoundText) & ")"
sql = sql & " GROUP BY dbo.TblPaymentType.Accountsus, dbo.TblTransactionPayments.PaymentID, dbo.TblPaymentType.PaymentName"
sql = sql & " ORDER BY dbo.TblTransactionPayments.PaymentID"


 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      
      For i = 1 To rs.RecordCount
      
       Notevalue = IIf(IsNull(rs("totals").value), 0, rs("totals").value)
      
                 
                  PaymentName = IIf(IsNull(rs("PaymentName").value), "", rs("PaymentName").value)
                 If Notevalue > 0 Then
                       StrTempAccountCode = IIf(IsNull(rs("Accountsus").value), "", rs("Accountsus").value)
                 '     PaymentName = IIf(IsNull(rs("PaymentName").value), "", rs("PaymentName").value)
                       If StrTempAccountCode = "" Then StrTempAccountCode = cashierBocaccount
                       If PaymentName = "" Then PaymentName = "‰ﬁœÌ…":     Notevalue = GetPointSAles(val(dcpoint.BoundText), 0, 2)
        
                                    
                              StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "  »‰«¡ ⁄·Ï „»Ì⁄«   " & PaymentName
                              LngDevNO = LngDevNO + 1
                  
                              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                  GoTo ErrTrap
                              End If
                 
                 ElseIf Notevalue < 0 Then
                 Notevalue = Abs(val(lblTotalsalecash.Caption))
                       StrTempAccountCode = IIf(IsNull(rs("Accountsus").value), "", rs("Accountsus").value)
                     ' PaymentName = IIf(IsNull(rs("PaymentName").value), "", rs("PaymentName").value)
                       If StrTempAccountCode = "" Then StrTempAccountCode = cashierBocaccount
                       If PaymentName = "" Then PaymentName = "‰ﬁœÌ…":     Notevalue = GetPointSAles(val(dcpoint.BoundText), 0, 2)
        
                                    
                              StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "  »‰«¡ ⁄·Ï „»Ì⁄«   " & PaymentName
                              LngDevNO = LngDevNO + 1
                  
                              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                                  GoTo ErrTrap
                              End If
                 
                 Else
                 
                End If
     
      rs.MoveNext
      Next i
      
    End If

    rs.Close
    
    
    
      Notevalue = Abs(val(lblTotalTransaction.Caption))
'«·œ«∆‰ «·„»Ì⁄« 
Account_Code_dynamic = get_account_code_branch(2, my_branch)
      StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "    «À»«  „»Ì⁄«   "
      
      StrTempAccountCode = Account_Code_dynamic
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            

End If


   '***********************«À»«  „»Ì⁄«  «·‰ﬁÿ…‰Â«ÌÂ  *****************************
   
   
'*************************„—œÊœ«  «·„»Ì⁄«  »œ«Ì…    *******************************
  Notevalue = val(Abs(lblReturn.Caption))
  Dim vcash As Double
'«·„œÌ‰ „—œÊœ«  «·„»Ì⁄« 
vcash = Abs(Me.lblTotalsalecash) - Abs(val(lblReturn.Caption))
StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "    «À»«  „—œÊœ«   "
If Notevalue >= 0 Then


Account_Code_dynamic = get_account_code_branch(3, my_branch)

      StrTempAccountCode = Account_Code_dynamic
   LngDevNO = LngDevNO + 1
            
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
            
           
    StrTempAccountCode = cashierBocaccount
    
If vcash < 0 Then
    
    
         LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(Me.lblTotalsalecash), 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
             StrTempAccountCode = DcboDebitSide.BoundText
        LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Abs(vcash), 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
    ElseIf vcash = 0 Then
    
    Else
      LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
    End If
    
 
            
            
            
 
            
End If

  '*************************„—œÊœ«  «·„»Ì⁄«  ‰Â«Ì…  *******************************
   
If val(Txtrecivedpetty.Text) > 0 Then ' ⁄Âœ… „”œœ…
    Notevalue = val(Txtrecivedpetty)
     StrTempAccountCode = DcboCreditSide.BoundText '⁄Âœ… «·„‘—›
    
             
            StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "   ”œ«œ „‰ «·⁄Âœ…  "
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If


      StrTempAccountCode = DcboDebitSide.BoundText  '⁄Âœ… «·ﬂ«‘Ì—
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            


      
            
            
End If


If val(TxtBalance.Text) > 0 Then   ' —’Ìœ ”«»ﬁ „”œœ
    Notevalue = val(TxtBalance)
              StrTempAccountCode = SupervisorBocaccount
            StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "   ”œ«œ „‰ «·—’Ìœ «·”«»ﬁ"
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If


      StrTempAccountCode = cashierBocaccount
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            

End If



If val(TxtPosSales.Text) > 0 Then   ' «Ã„«·Ì «·„»Ì⁄«  «·„”œœ
    Notevalue = val(TxtPosSales)
              StrTempAccountCode = SupervisorBocaccount
            StrTempDes = "”‰œ «€·«ﬁ ‰ﬁÿ… »Ì⁄ » «—ÌŒ  " & Now & "   ”œ«œ „‰   «·„»Ì⁄«  «·‰ﬁœÌ…"
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If


      StrTempAccountCode = cashierBocaccount
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            

End If


 
  
  
ErrTrap:
End Function

Private Sub DCboUserName_Change()
 Dim PettyId As Long
 Dim BoxID As Long
    If val(DCboUserName.BoundText) = 0 Then Me.DcboDebitSide.BoundText = "":   Exit Sub
    getCashireData val(DCboUserName.BoundText), PointID, Pointname, Balance, PettyId, pettyBalance, BoxID
Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", PettyId)
cashierBocaccount = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", BoxID)
    Me.LBLPOSCode = PointID
    Me.LBLPOSName = Pointname
    Me.LBLBalance.Caption = Balance
   LblPetty.Caption = pettyBalance
    Calc
End Sub

Private Sub DCboUserName_Click(Area As Integer)

DCboUserName_Change
End Sub

Function CheckAcconts() As Boolean
CheckAcconts = True
If Me.DcboDebitSide.BoundText = "" Then MsgBox "Õ”«» ⁄Âœ… «·ﬂ«‘Ì— €Ì— „Õœœ", vbCritical: CheckAcconts = False
If Me.DcboCreditSide.BoundText = "" Then MsgBox "Õ”«» ⁄Âœ… «·„‘—› €Ì— „Õœœ", vbCritical: CheckAcconts = False



         Account_Code_dynamic = get_account_code_branch(2, val(Current_branch))
        
            If Account_Code_dynamic = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "Branch Not Created", vbCritical
                        End If

              CheckAcconts = False
            ElseIf Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Account Not Defined in this Branch", vbCritical
                    End If
CheckAcconts = False
                
         
                End If
 
            
            
         Account_Code_dynamic1 = get_account_code_branch(3, val(Current_branch))
        
            If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "Branch Not Created", vbCritical
                        End If

              CheckAcconts = False
            ElseIf Account_Code_dynamic1 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» „ «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Account Not DefinœÊœ«  ed in this Branch", vbCritical
                    End If
CheckAcconts = False
                
         
                End If
                
 
End Function
Private Sub DCboUserName1_Change()

 Dim PettyId As Long
  Dim BoxID As Long
    If val(DCboUserName1.BoundText) = 0 Then Me.DcboCreditSide.BoundText = "":    Exit Sub
    getCashireData val(DCboUserName1.BoundText), , 0, , PettyId, , BoxID
Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", PettyId)
SupervisorBocaccount = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", BoxID)
Calc
End Sub

Private Sub DCboUserName1_Click(Area As Integer)
DCboUserName1_Change
End Sub

Private Sub dcpoint_Change()
'lblTotalTransaction.Caption = GetPointSAles(val(dcpoint.BoundText))
'Calc

End Sub
Function Calc()
LblTotals.Caption = val(LBLtTALCASHES.Caption) + val(LblPetty) + val(LBLBalance)
LblTotalCollected.Caption = val(TxtBalance) + val(Txtrecivedpetty) + val(TxtPosSales)
End Function
Private Sub dcpoint_Click(Area As Integer)
dcpoint_Change
End Sub

Private Sub dcShift_Click(Area As Integer)
   Dim ShiftFrom As Date
    Dim ShiftTo As Date
   GetShiftData val(dcShift.BoundText), ShiftFrom, ShiftTo

    LBLShiftID.Caption = val(dcShift.BoundText)
    ShfitFrom.value = ShiftFrom
    ShfitTo.value = ShiftTo
End Sub

Public Sub FillGridWithData(Optional PointID As Integer)

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
 
My_SQL = "SELECT     TOP 100 PERCENT SUM(dbo.TblTransactionPayments.[value] * ISNULL(dbo.TblTransactionPayments.Effect, 1)) AS totals, dbo.TblTransactionPayments.PaymentID, "
My_SQL = My_SQL & "                       dbo.TblPaymentType.PaymentName"
My_SQL = My_SQL & " FROM         dbo.TblTransactionPayments LEFT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
My_SQL = My_SQL & "      Where(dbo.TblTransactionPayments.locked Is Null  or dbo.TblTransactionPayments.locked=0) And (dbo.TblTransactionPayments.PointID = " & PointID & ")"
My_SQL = My_SQL & " GROUP BY dbo.TblTransactionPayments.PaymentID, dbo.TblPaymentType.PaymentName"
My_SQL = My_SQL & " ORDER BY dbo.TblTransactionPayments.PaymentID"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
lblTotalsalecash.Caption = 0
    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               
               If SystemOptions.UserInterface = ArabicInterface Then
                     If .TextMatrix(i, .ColIndex("PaymentName")) = "" Then .TextMatrix(i, .ColIndex("PaymentName")) = "‰ﬁœÌ "
               Else
               If .TextMatrix(i, .ColIndex("PaymentName")) = "" Then .TextMatrix(i, .ColIndex("PaymentName")) = "Cash "
               End If
               
              
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs.Fields("totals").value), 0, rs.Fields("totals").value)
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), IIf(IsNull(rs.Fields("totals").value), 0, rs.Fields("totals").value), rs.Fields("PaymentID").value)
             If i > 1 Then
               lblTotalsalecash.Caption = val(lblTotalsalecash.Caption) + .TextMatrix(i, .ColIndex("value"))
               
               End If
                rs.MoveNext
            Next
  lblTotalsalecash.Caption = val(lblTotalTransaction.Caption) - val(lblTotalsalecash.Caption)
            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Function print_ReportDay(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    XPDtbTrans.value = Date

    My_SQL = " SELECT     dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, "
    My_SQL = My_SQL & "                   dbo.TblItems.ItemNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.ItemDiscountType,"
    My_SQL = My_SQL & "                   dbo.Transaction_Details.ItemDiscount, dbo.Transaction_Details.Remarks, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,"
    My_SQL = My_SQL & "                   dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.QtyBySmalltUnit, dbo.Transaction_Details.TypeVAT,"
    My_SQL = My_SQL & "                   dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.Vat, dbo.Transaction_Details.Vatyo, dbo.Groups.GroupName,"
    My_SQL = My_SQL & "                   dbo.Groups.GroupNamee, dbo.Groups.Fullcode, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.barCodeNO, dbo.Transactions.Transaction_Type,"
    My_SQL = My_SQL & "                   dbo.TransactionTypes.TransactionTypeName , dbo.TransactionTypes.TransactionEnglishName, dbo.TblStore.StoreName, dbo.TblStore.storenamee ,dbo.TblItems.GroupID ,dbo.Groups.Separate"
    My_SQL = My_SQL & "        FROM         dbo.TblStore RIGHT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.Transaction_Details INNER JOIN"
    My_SQL = My_SQL & "                   dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON"
    My_SQL = My_SQL & "                   dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID"
    My_SQL = My_SQL & " where 1=1"
   My_SQL = My_SQL & " and  dbo.Transactions.UserID =" & user_id & ""
    My_SQL = My_SQL & " and  (dbo.Transactions.Transaction_Type = 9 or dbo.Transactions.Transaction_Type = 21)"
    My_SQL = My_SQL & " and  dbo.Transactions.Transaction_Date =" & SQLDate(XPDtbTrans.value, True) & ""
    My_SQL = My_SQL & " ORDER BY dbo.Transactions.Transaction_Type DESC"
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSalesDaylay.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSalesDaylay.rpt"
        End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
   
    CViewer.FireReport xReport, PrinterTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    My_SQL = "SELECT      SUM(dbo.TblTransactionPayments.[value] * ISNULL(dbo.TblTransactionPayments.Effect, 1)) AS totals, dbo.TblTransactionPayments.PaymentID, "
    My_SQL = My_SQL & "                       dbo.TblPaymentType.PaymentName"
    My_SQL = My_SQL & " FROM         dbo.TblTransactionPayments LEFT OUTER JOIN"
    My_SQL = My_SQL & "                       dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
    My_SQL = My_SQL & "      Where(dbo.TblTransactionPayments.locked Is Null  or dbo.TblTransactionPayments.locked=0) And (dbo.TblTransactionPayments.PointID = " & PointID & ")"
    My_SQL = My_SQL & " GROUP BY dbo.TblTransactionPayments.PaymentID, dbo.TblPaymentType.PaymentName"
    My_SQL = My_SQL & " ORDER BY dbo.TblTransactionPayments.PaymentID"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCloseCasher.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCloseCasher.rpt"
        End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
xReport.ParameterFields(4).AddCurrentValue dcpoint.Text
xReport.ParameterFields(5).AddCurrentValue DCboUserName1.Text
xReport.ParameterFields(6).AddCurrentValue LBLBalance.Caption
xReport.ParameterFields(7).AddCurrentValue TxtBalance.Text
xReport.ParameterFields(8).AddCurrentValue LblPetty.Caption
xReport.ParameterFields(9).AddCurrentValue Txtrecivedpetty.Text
xReport.ParameterFields(10).AddCurrentValue LBLtTALCASHES.Caption
xReport.ParameterFields(11).AddCurrentValue TxtPosSales.Text
xReport.ParameterFields(12).AddCurrentValue LblTotals.Caption
xReport.ParameterFields(13).AddCurrentValue LblTotalCollected.Caption
xReport.ParameterFields(14).AddCurrentValue DCboUserName.Text
xReport.ParameterFields(15).AddCurrentValue lblTotalTransaction.Caption
xReport.ParameterFields(16).AddCurrentValue lblReturn.Caption
xReport.ParameterFields(17).AddCurrentValue LblNet.Caption


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    'CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    
    CViewer.FireReport xReport, PrinterTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
'Function print_reportAll(Optional NoteSerial As String)
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String
'    DTPicker1.value = Date
'    My_SQL = "SELECT      SUM(dbo.TblTransactionPayments.[value] * ISNULL(dbo.TblTransactionPayments.Effect, 1)) AS totals, dbo.TblTransactionPayments.PaymentID, "
'    My_SQL = My_SQL & "                       dbo.TblPaymentType.PaymentName"
'    My_SQL = My_SQL & " FROM         dbo.TblTransactionPayments LEFT OUTER JOIN"
'    My_SQL = My_SQL & "                       dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
'    My_SQL = My_SQL & "      Where (dbo.TblTransactionPayments.PointID = " & PointID & ") and TblTransactionPayments.Recorddate=" & SQLDate(DTPicker1.value, True) & ""
'    My_SQL = My_SQL & " GROUP BY dbo.TblTransactionPayments.PaymentID, dbo.TblPaymentType.PaymentName"
'    My_SQL = My_SQL & " ORDER BY dbo.TblTransactionPayments.PaymentID"
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCloseCasherAll.rpt"
'        Else
''            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCloseCasherAll.rpt"
 '       End If
 '   If Dir(StrFileName) = "" Then
 '       'GetMsgs 139, vbExclamation
 '       Screen.MousePointer = vbDefault
 '       Exit Function
 '   End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
''    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'
'    Dim cCompanyInfo As New ClsCompanyInfo
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'        StrReportTitle = "" '& StrAccountName
'
'    Else
'
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
'        StrReportTitle = ""
'    End If
'
''    xReport.ParameterFields(3).AddCurrentValue user_name
'xReport.ParameterFields(4).AddCurrentValue dcpoint.Text
'xReport.ParameterFields(5).AddCurrentValue DCboUserName1.Text
'xReport.ParameterFields(6).AddCurrentValue LBLBalance.Caption
'xReport.ParameterFields(7).AddCurrentValue TxtBalance.Text
'xReport.ParameterFields(8).AddCurrentValue LblPetty.Caption
'xReport.ParameterFields(9).AddCurrentValue Txtrecivedpetty.Text
'xReport.ParameterFields(10).AddCurrentValue LBLtTALCASHES.Caption
'xReport.ParameterFields(11).AddCurrentValue TxtPosSales.Text
'xReport.ParameterFields(12).AddCurrentValue LblTotals.Caption
'xReport.ParameterFields(13).AddCurrentValue LblTotalCollected.Caption
'xReport.ParameterFields(14).AddCurrentValue DCboUserName.Text
'xReport.ParameterFields(15).AddCurrentValue lblTotalTransaction.Caption
'xReport.ParameterFields(16).AddCurrentValue lblReturn.Caption
'xReport.ParameterFields(17).AddCurrentValue LblNet.Caption
'xReport.ParameterFields(18).AddCurrentValue GetTimeTrans(GetMinTransID())
'xReport.ParameterFields(19).AddCurrentValue GetTimeTrans(GetMaxTransID())
'
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function
Function GetMinTransID(Optional EmpID As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT        MIN(Transaction_ID) AS ID"
sql = sql & " From dbo.transactions"
sql = sql & " WHERE        (Transaction_Date = " & SQLDate(DTPicker1.value, True) & ") and Emp_ID = " & EmpID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMinTransID = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetMinTransID = 0
End If
End Function
Function GetMaxTransID(Optional EmpID As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT        MAX(Transaction_ID) AS ID"
sql = sql & " From dbo.transactions"
sql = sql & " WHERE        (Transaction_Date = " & SQLDate(DTPicker1.value, True) & ") and Emp_ID = " & EmpID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxTransID = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetMaxTransID = 0
End If
End Function
Function GetTimeTrans(Optional Transaction_ID As Double) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT        TimeInvoice"
sql = sql & " From dbo.transactions"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetTimeTrans = IIf(IsNull(rs2("TimeInvoice").value), Time, rs2("TimeInvoice").value)
Else
GetTimeTrans = Time
End If
End Function
Private Sub Form_Load()
'    Resize_Form Me
    Dim My_SQL As String
    Dim Shiftcode As String
    Dim shiftname As String
    Dim FromDate1 As Date
    Dim ToDate1 As Date

Dim Dcombos As New ClsDataCombos
Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    DTPicker1.value = Date
        
        If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "select SeftCode,SheftName From  TbLSheft "
    Else
        My_SQL = "select SeftCode,SheftName From TbLSheft "
    End If

   fill_combo dcShift, My_SQL

If SystemOptions.HideInfroCasher = True Then
Frame4.Visible = False
Frame5.Visible = True
Frame6.Visible = True

End If

    'My_SQL = "select id,name From cachierData where ctype=0 "
    If SystemOptions.UserInterface = ArabicInterface Then
               My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.name"
               My_SQL = My_SQL & " FROM         dbo.cachierData LEFT OUTER JOIN"
               My_SQL = My_SQL & " dbo.TblShiftWorker ON dbo.cachierData.EmpID = dbo.TblShiftWorker.EmpID"
            
               My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 0)"
    Else
                   My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.namee"
               My_SQL = My_SQL & " FROM         dbo.cachierData LEFT OUTER JOIN"
               My_SQL = My_SQL & " dbo.TblShiftWorker ON dbo.cachierData.EmpID = dbo.TblShiftWorker.EmpID"
            
               My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 0)"

    
    End If
    
      My_SQL = My_SQL & " and  PointId in (SELECT     BoxID  From dbo.Tblposdata  Where (BranchId = " & Current_branch & ")) "
My_SQL = My_SQL & "       and (isCachDeactivated is null  or isCachDeactivated=0)"
    fill_combo DCboUserName, My_SQL
        If SystemOptions.UserInterface = ArabicInterface Then

        My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.name"
    My_SQL = My_SQL & " FROM         dbo.cachierData  "
   My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 1)"
    Else
    
        My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.namee"
    My_SQL = My_SQL & " FROM         dbo.cachierData  "
   My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 1)"
    End If
      My_SQL = My_SQL & " and  PointId in (SELECT     BoxID  From dbo.Tblposdata  Where (BranchId = " & Current_branch & ")) "
    
    My_SQL = My_SQL & "       and (isCachDeactivated is null  or isCachDeactivated=0)"
        fill_combo DCboUserName1, My_SQL
        
        
            If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "select BoxID,BoxName From Tblposdata    where BranchId =" & Current_branch
    Else
        My_SQL = "select BoxID,BoxNamee From Tblposdata   where BranchId =" & Current_branch
    End If

    fill_combo dcpoint, My_SQL

   dcpoint.BoundText = PPointID
DCboUserName.BoundText = CurrentCashireID
If ScreenJob = 1 Then
Me.Caption = "«€·«ﬁ «·‰ﬁÿ…"
LblHeader.Caption = Me.Caption
 End If

 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
End Sub

Private Sub Image2_Click()
Call Shell("OSK.exe")
End Sub

Private Sub TxtBalance_Change()
Calc
End Sub

Private Sub TxtBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtBalance.Text, 0)
End Sub

Private Sub TxtPosSales_Change()
Calc
End Sub

Private Sub Txtrecivedpetty_Change()
Calc
End Sub
