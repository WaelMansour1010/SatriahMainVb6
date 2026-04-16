VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLocedkOrderMaintAqarSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·»«  «ﬁ›«· «·’Ì«‰Â"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15180
   Icon            =   "FrmLockedOrderMaintAqarSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «ﬁ›«· «·ÿ·»"
      Height          =   1035
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   3600
      Width           =   3975
      Begin MSComCtl2.DTPicker LocTo 
         Height          =   330
         Left            =   1770
         TabIndex        =   47
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker LocFrom 
         Height          =   330
         Left            =   1770
         TabIndex        =   46
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal LocFromH 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal LocToH 
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   16
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   15
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   540
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «‰ Â«¡ «·ÿ·»"
      Height          =   1035
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   2640
      Width           =   3975
      Begin MSComCtl2.DTPicker EndFrom 
         Height          =   330
         Left            =   1770
         TabIndex        =   40
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker EndTo 
         Height          =   330
         Left            =   1770
         TabIndex        =   39
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal EndFromH 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal EndToH 
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   14
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   12
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ »œ«Ì… «·ÿ·»"
      Height          =   1035
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   2640
      Width           =   3975
      Begin MSComCtl2.DTPicker SatrTO 
         Height          =   330
         Left            =   1770
         TabIndex        =   33
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker SatrFrom 
         Height          =   330
         Left            =   1770
         TabIndex        =   32
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal SatrFromH 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal SatrTOH 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   11
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   9
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1965
      Index           =   1
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3600
      Width           =   10755
      Begin VB.TextBox TxtRemark 
         Alignment       =   1  'Right Justify
         Height          =   1695
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   120
         Width           =   3255
      End
      Begin VB.CheckBox ChLock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» «ﬁ›«· «·ÿ·» "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox DcbStatus 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmLockedOrderMaintAqarSearch.frx":038A
         Left            =   8640
         List            =   "FrmLockedOrderMaintAqarSearch.frx":0394
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtOrderWork 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox TxtSearchCodeSuper 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DcbIqara 
         Height          =   315
         Left            =   3960
         TabIndex        =   20
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpNameSuper 
         Height          =   315
         Left            =   3960
         TabIndex        =   21
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   3960
         TabIndex        =   25
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
         Height          =   435
         Index           =   18
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ«·… «·’Ì«‰Â"
         Height          =   375
         Index           =   17
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„· «·„ÿ·Ê»"
         Height          =   435
         Index           =   8
         Left            =   9690
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ÊŸ› «·’Ì«‰Â"
         Height          =   285
         Index           =   7
         Left            =   9570
         TabIndex        =   26
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   255
         Index           =   13
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”ƒÊ· «·’Ì«‰Â"
         Height          =   285
         Index           =   0
         Left            =   9570
         TabIndex        =   22
         Top             =   495
         Width           =   1125
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ  «·ÿ·»"
      Height          =   1035
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   1770
         TabIndex        =   6
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   1770
         TabIndex        =   7
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateFromh 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToh 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   885
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2640
      Width           =   2835
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15075
      _cx             =   26591
      _cy             =   4630
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
      Rows            =   50
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmLockedOrderMaintAqarSearch.frx":03A4
      ScrollTrack     =   -1  'True
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   11
      Top             =   5040
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   12
      Top             =   5040
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   5040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4740
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   8280
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   2775
   End
End
Attribute VB_Name = "FrmLocedkOrderMaintAqarSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public m_RetrunType As Integer

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       
 GetData
         
        Case 1
            clear_all Me
            LocFrom.value = ""
            LocTo.value = ""
DtpDateFrom.value = ""
DtpDateTo.value = ""
SatrFrom.value = ""
SatrTO.value = ""
EndFrom.value = ""
EndTo.value = ""

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
End Sub

Private Sub DtpDateFrom_Change()
 DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
End Sub



Private Sub DtpDateFromH_LostFocus()
        VBA.Calendar = vbCalGreg
           DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub

Private Sub DtpDateTo_Change()
 DtpDateToH.value = ToHijriDate(DtpDateTo.value)
End Sub



Private Sub DtpDateToH_LostFocus()
   VBA.Calendar = vbCalGreg
           DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub EndFrom_Change()
EndFromH.value = ToHijriDate(EndFrom.value)
End Sub

Private Sub EndFromH_LostFocus()
 VBA.Calendar = vbCalGreg
           EndFrom.value = ToGregorianDate(EndFromH.value)
End Sub

Private Sub EndTo_Change()
EndToH.value = ToHijriDate(EndTo.value)
End Sub

Private Sub EndToH_LostFocus()
 VBA.Calendar = vbCalGreg
           EndTo.value = ToGregorianDate(EndToH.value)
End Sub

Private Sub Fg_Click()

    With Me.fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("id"))) = 0 Then
            Exit Sub
        End If

      
               FrmLockedOrderMaintenance.Retrive val(.TextMatrix(.Row, .ColIndex("id")))
                
      

    End With

End Sub


 

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcboEmpNameSuper
  

    Dcombos.GetIqar DcbIqara

    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
      SetDtpickerDate Me.SatrFrom
    SetDtpickerDate Me.SatrTO
      SetDtpickerDate Me.EndFrom
    SetDtpickerDate Me.EndTo
      SetDtpickerDate Me.LocFrom
    SetDtpickerDate Me.LocTo
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = " SELECT     dbo.TblLockedOrderMaintenance.ID, dbo.TblLockedOrderMaintenance.RecDateH, dbo.TblLockedOrderMaintenance.RecDate, "
StrSQL = StrSQL & "                      dbo.TblLockedOrderMaintenance.OrderNo, dbo.TblLockedOrderMaintenance.LocationIqar, dbo.TblLockedOrderMaintenance.Remark,"
StrSQL = StrSQL & "                       dbo.TblLockedOrderMaintenance.OrderWork, dbo.TblLockedOrderMaintenance.DMY, dbo.TblLockedOrderMaintenance.Status, dbo.TblLockedOrderMaintenance.Cont,"
StrSQL = StrSQL & "                       dbo.TblLockedOrderMaintenance.EndFateH, dbo.TblLockedOrderMaintenance.EndFate, dbo.TblLockedOrderMaintenance.Lock,"
StrSQL = StrSQL & "                       dbo.TblLockedOrderMaintenance.SatarDateH, dbo.TblLockedOrderMaintenance.SatarDate, dbo.TblLockedOrderMaintenance.LockDateH,"
StrSQL = StrSQL & "                       dbo.TblLockedOrderMaintenance.LockDate, dbo.TblLockedOrderMaintenance.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                       dbo.TblLockedOrderMaintenance.AqrID, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblLockedOrderMaintenance.EmpID, TblEmployee_1.Emp_Code,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4,"
StrSQL = StrSQL & "                       TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee4, TblEmployee_1.Emp_Namee3, TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee1,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee, dbo.TblLockedOrderMaintenance.SuperVM, TblEmployee_1.Emp_Code AS Emp_CodeSup,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name AS Emp_NameSup, TblEmployee_1.Emp_Name1 AS Emp_Name1Sup, TblEmployee_1.Emp_Name2 AS Emp_Name2Sup,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name3 AS Emp_Name3Sup, TblEmployee_1.Emp_Name4 AS Emp_Name4Sup, TblEmployee_1.Fullcode AS FullcodeSup,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee4 AS Emp_Namee4Sup, TblEmployee_1.Emp_Namee3 AS Emp_Namee3Sup, TblEmployee_1.Emp_Namee2 AS Emp_Namee2Sup,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee1 AS Emp_Namee1Sup, TblEmployee_1.Emp_Namee AS Emp_NameeSup"
StrSQL = StrSQL & "  FROM         dbo.TblLockedOrderMaintenance LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblLockedOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_2 ON dbo.TblLockedOrderMaintenance.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblAqar ON dbo.TblLockedOrderMaintenance.AqrID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.TblLockedOrderMaintenance.BranchID = dbo.TblBranchesData.branch_id"

  StrWhere = ""
    BolBegine = False
    

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.ID >=" & val(Me.TxtIDFrom.Text) & ""
       Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
   

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
           BolBegine = True
           StrWhere = " Where dbo.TblLockedOrderMaintenance.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    '///////////////////
   
'////////////////////////
 If Me.ChLock.value = vbChecked Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.Lock = 1 "
        Else
            BolBegine = True
              StrWhere = " Where dbo.TblLockedOrderMaintenance.Lock = 1 "
        End If
    End If
    
 If val(Me.DcbStatus.ListIndex) <> -1 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.Status = " & val(Me.DcbStatus.ListIndex) & ""
        Else
            BolBegine = True
              StrWhere = " Where dbo.TblLockedOrderMaintenance.Status = " & val(Me.DcbStatus.ListIndex) & ""
        End If
    End If
    
 If Me.txtRemark.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.Remark like '%" & Me.txtRemark.Text & "%'"
        Else
            BolBegine = True
           StrWhere = " Where dbo.TblLockedOrderMaintenance.Remark like '%" & Me.txtRemark.Text & "%'"
        End If
    End If
    
 If Me.TxtOrderWork.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.OrderWorklike '%" & Me.TxtOrderWork.Text & "%'"
        Else
            BolBegine = True
           StrWhere = " Where dbo.TblLockedOrderMaintenance.OrderWork like '%" & Me.TxtOrderWork.Text & "%'"
        End If
    End If
  If Me.DcboEmpNameSuper.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.SuperVM=" & Me.DcboEmpNameSuper.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.SuperVM=" & Me.DcboEmpNameSuper.BoundText & ""
        End If
    End If
       If Me.DcboEmpName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.EmpID=" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.EmpID=" & Me.DcboEmpName.BoundText & ""
        End If
    End If
   If Me.DcbIqara.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid=" & Me.DcbIqara.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAqar.Aqarid=" & Me.DcbIqara.BoundText & ""
        End If
    End If
   
   If Not IsNull(Me.LocFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.LockDate >=" & SQLDate(Me.LocFrom.value, True) & ""
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.Lock = 1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.LockDate >=" & SQLDate(Me.LocFrom.value, True) & ""
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.Lock = 1 "
        End If
    End If

    If Not IsNull(Me.LocTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLockedOrderMaintenance.LockDate <=" & SQLDate(Me.LocTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLockedOrderMaintenance.LockDate <=" & SQLDate(Me.LocTo.value, True) & ""
        End If
    End If
    
   
   
  If Not IsNull(Me.EndFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.EndFate >=" & SQLDate(Me.EndFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.EndFate >=" & SQLDate(Me.EndFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.EndTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLockedOrderMaintenance.EndFate <=" & SQLDate(Me.EndTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLockedOrderMaintenance.EndFate <=" & SQLDate(Me.EndTo.value, True) & ""
        End If
    End If
    
    

    If Not IsNull(Me.SatrFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.SatarDate >=" & SQLDate(Me.SatrFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.SatarDate >=" & SQLDate(Me.SatrFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.SatrTO.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLockedOrderMaintenance.SatarDate <=" & SQLDate(Me.SatrTO.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLockedOrderMaintenance.SatarDate <=" & SQLDate(Me.SatrTO.value, True) & ""
        End If
    End If

  If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLockedOrderMaintenance.RecDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLockedOrderMaintenance.RecDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLockedOrderMaintenance.RecDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLockedOrderMaintenance.RecDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
    
    
    
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblLockedOrderMaintenance.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                    .TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", rs("aqarname").value)
                If Not (IsNull(rs("RecDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecDate")) = Format(rs("RecDate").value, "yyyy/M/d")
                End If
                 .TextMatrix(i, .ColIndex("RecDateH")) = IIf(IsNull(rs("RecDateH").value), "", rs("RecDateH").value)
                 If Not (IsNull(rs("SatarDate").value)) Then
                    .TextMatrix(i, .ColIndex("SatarDate")) = Format(rs("SatarDate").value, "yyyy/M/d")
                End If
                 .TextMatrix(i, .ColIndex("SatarDateH")) = IIf(IsNull(rs("SatarDateH").value), "", rs("SatarDateH").value)
                 If rs("Lock").value = True Then
                  If Not (IsNull(rs("LockDate").value)) Then
                    .TextMatrix(i, .ColIndex("LockDate")) = Format(rs("LockDate").value, "yyyy/M/d")
                End If
                 .TextMatrix(i, .ColIndex("LockDateH")) = IIf(IsNull(rs("LockDateH").value), "", rs("LockDateH").value)
                  .TextMatrix(i, .ColIndex("Lock")) = -1
                 Else
                 .TextMatrix(i, .ColIndex("LockDateH")) = ""
                 .TextMatrix(i, .ColIndex("LockDate")) = ""
                 .TextMatrix(i, .ColIndex("Lock")) = 0
                 End If
                 
                  If Not (IsNull(rs("EndFate").value)) Then
                    .TextMatrix(i, .ColIndex("EndFate")) = Format(rs("EndFate").value, "yyyy/M/d")
                End If
                 .TextMatrix(i, .ColIndex("EndFateH")) = IIf(IsNull(rs("EndFateH").value), "", rs("EndFateH").value)
                 
                .TextMatrix(i, .ColIndex("OrderWork")) = IIf(IsNull(rs("OrderWork").value), "", rs("OrderWork").value)
                 
             .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
            If val(rs("Status").value) = 0 Then
            .TextMatrix(i, .ColIndex("Status")) = "·„ Ì „"
            Else
             .TextMatrix(i, .ColIndex("Status")) = " „"
            End If
            
                If SystemOptions.UserInterface = EnglishInterface Then
              .TextMatrix(i, .ColIndex("Emp_NameSup")) = IIf(IsNull(rs("Emp_NameeSup").value), "", rs("Emp_NameeSup").value)
          
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                Else
               .TextMatrix(i, .ColIndex("Emp_NameSup")) = IIf(IsNull(rs("Emp_NameSup").value), "", rs("Emp_NameSup").value)
               .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
    End If
              
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
  '
End Sub

Private Sub LocFrom_Change()
LocFromH.value = ToHijriDate(LocFrom.value)
End Sub

Private Sub LocFromH_LostFocus()
VBA.Calendar = vbCalGreg
           LocFrom.value = ToGregorianDate(LocFromH.value)
End Sub

Private Sub LocTo_Change()
LocToH.value = ToHijriDate(LocTo.value)
End Sub

Private Sub LocToH_LostFocus()
VBA.Calendar = vbCalGreg
           LocTo.value = ToGregorianDate(LocToH.value)
End Sub

Private Sub SatrFrom_Change()
SatrFromH.value = ToHijriDate(SatrFrom.value)
End Sub

Private Sub SatrFromH_LostFocus()
 VBA.Calendar = vbCalGreg
           SatrFrom.value = ToGregorianDate(SatrFromH.value)
End Sub

Private Sub SatrTO_Change()
 SatrTOH.value = ToHijriDate(SatrTO.value)
End Sub

Private Sub SatrTOH_LostFocus()
 VBA.Calendar = vbCalGreg
           SatrTO.value = ToGregorianDate(SatrTOH.value)
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmAqarSearch.m_RetrunType = 2020
FrmAqarSearch.show


End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub

Private Sub DcboEmpNameSuper_Change()
DcboEmpNameSuper_Click (0)
End Sub

Private Sub DcboEmpNameSuper_Click(Area As Integer)

   If val(DcboEmpNameSuper.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpNameSuper.BoundText, EmpCode
    TxtSearchCodeSuper.Text = EmpCode
End Sub
Private Sub TxtSearchCodeSuper_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCodeSuper.Text, EmpID
        DcboEmpNameSuper.BoundText = EmpID
    End If
End Sub
Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

