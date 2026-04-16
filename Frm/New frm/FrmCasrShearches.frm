VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCasrShearches 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·„—ﬂ»« "
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14655
   Icon            =   "FrmCasrShearches.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   14655
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
   Begin VB.TextBox txtOperatorN 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   11040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   4920
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12450
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   3480
      Width           =   1065
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   11040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   2760
      Width           =   2475
   End
   Begin VB.TextBox TxtNotes 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   4320
      Width           =   5115
   End
   Begin VB.TextBox txtBoardNO 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   11040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3120
      Width           =   2475
   End
   Begin VB.TextBox TxtLicenseNO 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   11040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3840
      Width           =   2475
   End
   Begin VB.TextBox txtModel 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   11040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4200
      Width           =   2475
   End
   Begin VB.TextBox txtLastKMCounter 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   6720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3840
      Width           =   2955
   End
   Begin VB.TextBox VehicleLong 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   11040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4560
      Width           =   2475
   End
   Begin VB.TextBox TxtEquQty 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2760
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3480
      Width           =   2475
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·‘—«¡"
      Height          =   1035
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3150
      Width           =   2415
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94699523
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   630
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94699523
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1815
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   330
         Width           =   660
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   14700
      _cx             =   25929
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
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCasrShearches.frx":038A
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
      Left            =   5880
      TabIndex        =   6
      Top             =   4920
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
      Left            =   5010
      TabIndex        =   7
      Top             =   4920
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
      Left            =   4230
      TabIndex        =   8
      Top             =   4920
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
   Begin MSDataListLib.DataCombo DCGroup 
      Height          =   315
      Left            =   2760
      TabIndex        =   12
      Top             =   2760
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcEmployee 
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   3480
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcFixedAssets 
      Height          =   315
      Left            =   6720
      TabIndex        =   14
      Tag             =   "Õœœ «”„ «·„⁄œ…"
      Top             =   2760
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCInsuranceCompanyId 
      Height          =   315
      Left            =   2760
      TabIndex        =   18
      Top             =   3120
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Dynamic_Byte.NourHijriCal DpLicenseExpireDateH 
      Height          =   315
      Left            =   6720
      TabIndex        =   34
      Top             =   4560
      Width           =   2835
      _ExtentX        =   3413
      _ExtentY        =   556
   End
   Begin Dynamic_Byte.NourHijriCal DpTestExpireDateH 
      Height          =   315
      Left            =   2760
      TabIndex        =   36
      Top             =   3840
      Width           =   2475
      _ExtentX        =   4577
      _ExtentY        =   556
   End
   Begin Dynamic_Byte.NourHijriCal DpInsuranceExpireDateH 
      Height          =   315
      Left            =   6720
      TabIndex        =   37
      Top             =   4200
      Width           =   2835
      _ExtentX        =   4577
      _ExtentY        =   556
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic7 
      Height          =   435
      Left            =   6720
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3120
      Width           =   4365
      _cx             =   7699
      _cy             =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.TextBox txtLetter1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3630
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   0
         Width           =   540
      End
      Begin VB.TextBox txtLetter2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3210
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   0
         Width           =   450
      End
      Begin VB.TextBox txtLetter3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2700
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtNum1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1500
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   0
         Width           =   675
      End
      Begin VB.TextBox txtNum2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   900
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox txtNum3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   510
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   0
         Width           =   555
      End
      Begin VB.TextBox txtLetter4 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2175
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   0
         Width           =   675
      End
      Begin VB.TextBox txtNum4 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   0
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
      Height          =   315
      Index           =   5
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   4920
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„⁄œÂ/«·”Ì«—…"
      Height          =   255
      Index           =   0
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ  ‰Â«Ì… «· √„Ì‰"
      Height          =   315
      Index           =   127
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   4320
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì… «·›Õ’"
      Height          =   375
      Index           =   120
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì… «·«” „«—…"
      Height          =   255
      Index           =   128
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê’› «·„⁄œÂ/«·”Ì«—…"
      Height          =   195
      Index           =   124
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4440
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊœÌ·"
      Height          =   315
      Index           =   107
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «··ÊÕ…"
      Height          =   255
      Index           =   105
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·«” „«—…"
      Height          =   315
      Index           =   106
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3840
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Œ— ﬁ—«¡… ··⁄œ«œ"
      Height          =   255
      Index           =   11
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3960
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿÊ· «·„—ﬂ»…"
      Height          =   315
      Index           =   9
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ„Ê·…"
      Height          =   315
      Index           =   8
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‘—ﬂ… «· √„Ì‰"
      Height          =   375
      Index           =   7
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁ«∆œ «·„⁄œÂ/«·”Ì«—…"
      Height          =   315
      Index           =   104
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
      Height          =   315
      Index           =   103
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„⁄œÂ/«·”Ì«—…"
      Height          =   315
      Index           =   102
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1650
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2820
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
      TabIndex        =   10
      Top             =   2820
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   435
      Index           =   10
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2700
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCasrShearches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch
Public SendForm As String


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            GetData

        Case 1
            clear_all Me
DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
    DpInsuranceExpireDateH.value = "1020/02/02"
 DpLicenseExpireDateH.value = "1020/02/02"
  DpTestExpireDateH.value = "1020/02/02"
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub





Private Sub DcEmployee_Change()
DcEmployee_Click (0)
End Sub

Private Sub DcEmployee_Click(Area As Integer)

       If val(DcEmployee.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcEmployee.BoundText, EmpCode
    Text2.Text = EmpCode
End Sub

 

Private Sub Fg_Click()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
            Exit Sub
        End If
            
           If SendForm = "VA" Then
           
                FrmVehicleAllocation.Add_Board val(.TextMatrix(.Row, .ColIndex("id"))), FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.Col
                FrmVehicleAllocation.FgInstallments.TextMatrix(FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.ColIndex("BoardNo")) = (.TextMatrix(.Row, .ColIndex("BoardNO")))
              
               ' FrmVehicleAllocation.FgInstallments.TextMatrix(FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.ColIndex("CarID")) = val(.TextMatrix(.Row, .ColIndex("id")))
               ' FrmVehicleAllocation.FgInstallments.TextMatrix(FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.ColIndex("BoardNo")) = (.TextMatrix(.Row, .ColIndex("BoardNO")))
              '
              '  FrmVehicleAllocation.FgInstallments.TextMatrix(FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.ColIndex("carstudentcount")) = (.TextMatrix(.Row, .ColIndex("capacity")))
              ''  FrmVehicleAllocation.FgInstallments.TextMatrix(FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.ColIndex("DriverID")) = (.TextMatrix(.Row, .ColIndex("Emp_ID")))
              ''  FrmVehicleAllocation.FgInstallments.TextMatrix(FrmVehicleAllocation.FgInstallments.Row, FrmVehicleAllocation.FgInstallments.ColIndex("Driver")) = (.TextMatrix(.Row, .ColIndex("Emp_Name")))
              '  FrmVehicleAllocation.FgInstallments.Rows = FrmVehicleAllocation.FgInstallments.Rows + 1
            
           
            ElseIf SendForm = "FrmOut" Then
                 ' FrmOut.Text6.Text = .TextMatrix(.Row, .ColIndex("FullCode"))
                  FrmOut.DCEquipments.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
                ElseIf SendForm = "OrderUpload" Then
                  FrmOrderUpload.DcbCar.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))
                ElseIf SendForm = "TravelTrans" Then
                  FrmTravelTransactions.DCCar.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))

            ElseIf SendForm = "FrmMantinanceReport" Then
                  FrmMantinanceReport.DcbEqup.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
             ElseIf SendForm = "FrmMantinanceReport2" Then
             FrmMantinanceReport.DcbEqup2.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
               ElseIf SendForm = "FrmCarsPlan" Then
                   FrmCarsPlan.DCCar.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))

            ElseIf SendForm = "DriverAllocation" Then
             '   FrmDriverAllocation.dcCars.SetFocus
                    FrmDriverAllocation.txtcarcode.Text = .TextMatrix(.Row, .ColIndex("FullCode"))
                    FrmDriverAllocation.dcCars.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))
              ElseIf SendForm = "OrderMaintin" Then
              FrmOrderMaintin.DcbEquepment.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
             FrmOrderMaintin.DcbEquepment_Change
            ElseIf SendForm = "RequerMainten" Then
              FrmRequerMainten.DcbEquepment.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
               FrmRequerMainten.DcbEquepment_Change
             ElseIf SendForm = "MovingEmp" Then
              FrmMovingEmp2.DcbEquepment.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
              FrmMovingEmp2.DcbEquepment_Change
             ElseIf SendForm = "frmdriveassestMove" Then
                 Dim AsID As Double
                     frmdriveassestMove.GetCardID val(.TextMatrix(.Row, .ColIndex("FixedassetId"))), AsID, 1
          frmdriveassestMove.dcmboassest.BoundText = AsID
          frmdriveassestMove.dcmboassest_Change
         '   frmdriveassestMove.dcmboassest.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))


            ElseIf SendForm = "FrmAccidentReport" Then
            FrmAccidentReport.TxtPlateNo.Text = (.TextMatrix(.Row, .ColIndex("BoardNO")))
        
        ElseIf SendForm = "TravelRports" Then
            frmTravelRports.DcbCar.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))
          ElseIf SendForm = "FrmAccidentReport2" Then
                       FrmAccidentReport.DcbCrID.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))
            frmdriveassestMove.dcmboassest.BoundText = val(.TextMatrix(.Row, .ColIndex("FixedassetId")))
            
                    ElseIf SendForm = "FrmSearchinvestment" Then
                       FrmSearchinvestment.DCCar.BoundText = val(.TextMatrix(.Row, .ColIndex("id")))
             
            
            
           Else
              FrmCars.Retrive val(.TextMatrix(.Row, .ColIndex("id")))
       End If
    

    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.Text, EmpID
        DcEmployee.BoundText = EmpID
    End If
End Sub

Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub
Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
Cal_Board

End Sub
Private Sub Cal_Board()
    txtBoardNO.Text = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text
End Sub
Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board

End Sub
Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub ChangeLang()
    'Dim XPic As IPictureDisp
    'Set XPic = Me.XPBtnMove(1).ButtonImage
    'Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
    'Set XPic = Me.XPBtnMove(0).ButtonImage
    'Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    'Set Me.XPBtnMove(3).ButtonImage = XPic
'    Label1.Visible = False

    'Cmd(0).Caption = "New"
    'Cmd(1).Caption = "Edit"
    'Cmd(2).Caption = "Save"
    'Cmd(3).Caption = "Undo"
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
 'Cmd(9).Caption = "Print"
    Cmd(2).Caption = "Exit"
 '

    Me.Caption = " Search For Dataof Cars    "
    'EleHeader.Caption = Me.Caption
       Me.lbl(0).Caption = "Code"
    Me.lbl(102).Caption = "Name"

    Me.lbl(103).Caption = "Type"
   ' Me.lbl(117).Caption = "Branch"
    Me.lbl(104).Caption = "Employee"
    Me.lbl(7).Caption = "Insur. Co."
    Me.lbl(105).Caption = "Board No."
    Me.Fra(1).Caption = "Purchase Date"
    Me.lbl(120).Caption = "Check Up Date"
    Me.lbl(106).Caption = "License No."
    Me.lbl(102).Caption = "Name"
    Me.lbl(128).Caption = "License Expire"
    Me.lbl(127).Caption = "Insurance Expire"

    Me.lbl(107).Caption = "Model"
    Me.lbl(11).Caption = "Last Km Count"
    Me.lbl(124).Caption = "Remarks"
        Me.lbl(8).Caption = "loader"
    Me.lbl(9).Caption = "VehicleLong"
    'Cmd(10).Caption = "Maintenance Plan"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(2).Caption = "Total"
   'lbl(8).Caption = "By"
   ' lbl(7).Caption = "Curr rec."
   ' lbl(6).Caption = "rec. count"

   With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "NO"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
         .TextMatrix(0, .ColIndex("FixedAssets_Name")) = "Name"
        .TextMatrix(0, .ColIndex("TBLCarTypes_name")) = "Type"
         .TextMatrix(0, .ColIndex("Model1")) = "Model"
        .TextMatrix(0, .ColIndex("VehicleLong")) = "VehicleLong"
        .TextMatrix(0, .ColIndex("EquQty")) = "loader"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
        
        .TextMatrix(0, .ColIndex("BoardNO")) = "Board No."
        .TextMatrix(0, .ColIndex("LastKMCounter")) = "Last Km Count"
        .TextMatrix(0, .ColIndex("Inst_Name")) = "Insur. Co."
         .TextMatrix(0, .ColIndex("LicenseNO")) = "License No."
        .TextMatrix(0, .ColIndex("PurchaseDate")) = "Purchase Date"
         .TextMatrix(0, .ColIndex("LicenseExpireDateH")) = "License Expire"
        .TextMatrix(0, .ColIndex("InsuranceExpireDateH")) = "Insurance Expire"
        .TextMatrix(0, .ColIndex("TestExpireDateH")) = "Check Up Date"
        .TextMatrix(0, .ColIndex("Notes")) = "Remarks"
        

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetItemsNames DCItem, , , , True
   ' Set DCboSearch = New clsDCboSearch
   ' Set DCboSearch.Client = Me.DCItem
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

   Dcombos.GetTblCarsDataGroup DCGroup
  Dcombos.GetFixedAssets Me.DcFixedAssets, True
  Dim My_SQL As String
    My_SQL = "  select id,name from insurance_companies   "
    fill_combo DCInsuranceCompanyId, My_SQL
       If SystemOptions.UserInterface = ArabicInterface Then
         My_SQL = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
     Else
         My_SQL = "  select   e.Emp_ID Emp_ID , e.Emp_NameE   Emp_NameE  from TblEmployee e, TblEmpJobsTypes  j"
     End If
         My_SQL = My_SQL & "   Where e.JobTypeID = j.JobTypeID"
         My_SQL = My_SQL & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
    fill_combo DcEmployee, My_SQL
    

   ' Dcombos.GetBranches dcBranch
    
 ' Dcombos.GetEmployees Me.DcEmployee, , True
 
     DpInsuranceExpireDateH.value = "1020/02/02"
 DpLicenseExpireDateH.value = "1020/02/02"
  DpTestExpireDateH.value = "1020/02/02"
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me
    Fg.ColHidden(5) = True

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    
    'SendForm = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
    SendForm = ""
End Sub

Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 
  
StrSQL = " SELECT     dbo.TblCarsData.id, dbo.TblCarsData.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCarsData.code, "
StrSQL = StrSQL & "                      dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix, dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, dbo.TblCarsData.fixedAssetid,"
StrSQL = StrSQL & "                       dbo.FixedAssets.id AS FixedAssets_ID, dbo.FixedAssets.Name AS FixedAssets_Name, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name AS TBLCarTypes_name,"
StrSQL = StrSQL & "                       dbo.TBLCarTypes.namee, dbo.TblCarsData.Emp_id, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Emp_FullCode,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Namee, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.VehicleLong, dbo.TblCarsData.PurchaseDate,"
StrSQL = StrSQL & "                       dbo.TblCarsData.LastKMCounter, dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.EquQty, dbo.TblCarsData.Notes, dbo.TblCarsData.InsuranceCompanyId,"
StrSQL = StrSQL & "                       dbo.insurance_companies.name AS Inst_Name, dbo.insurance_companies.Namee AS Inst_NameE, dbo.TblCarsData.TestExpireDateH,"
StrSQL = StrSQL & "                       dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.LicenseExpireDateH, dbo.TblCarsData.TestExpireDate, dbo.TblCarsData.InsuranceExpireDate,"
StrSQL = StrSQL & "                       dbo.FixedAssets.namee AS FixedAssets_NameE , dbo.TblCarsData.Capacity , dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile,dbo.TblCarsData.OperatorN"
StrSQL = StrSQL & "  FROM         dbo.TblCarsData LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in(" & Current_branchSql & "))"
 'where 1=1
 
    BolBegine = True
    StrWhere = ""
    
        If Me.DpTestExpireDateH.value <> "1020/02/02" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.TestExpireDateH='" & Me.DpTestExpireDateH.value & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.TestExpireDateH='" & Me.DpTestExpireDateH.value & "'"
        End If
    End If
    
       If Me.DpLicenseExpireDateH.value <> "1020/02/02" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.LicenseExpireDateH='" & Me.DpLicenseExpireDateH.value & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.LicenseExpireDateH='" & Me.DpLicenseExpireDateH.value & "'"
        End If
    End If
    
       If Me.DpInsuranceExpireDateH.value <> "1020/02/02" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.InsuranceExpireDateH='" & Me.DpInsuranceExpireDateH.value & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.InsuranceExpireDateH='" & Me.DpInsuranceExpireDateH.value & "'"
        End If
    End If
    
     If Me.TxtEquQty.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.EquQty=" & Me.TxtEquQty.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.EquQty=" & Me.TxtEquQty.Text & ""
        End If
    End If
         If Me.TxtOperatorN.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.OperatorN=N'" & TxtOperatorN.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.OperatorN=N'" & TxtOperatorN.Text & "'"
        End If
    End If
    
 If Me.VehicleLong.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.VehicleLong=" & Me.VehicleLong.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.VehicleLong=" & Me.VehicleLong.Text & ""
        End If
    End If
        If Me.TxtNotes.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.Notes like '%" & Me.TxtNotes.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.Notes  like '%" & Me.TxtNotes.Text & "%'"
        End If
    End If
    
     If Me.txtLastKMCounter.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.LastKMCounter='" & Me.txtLastKMCounter.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.LastKMCounter='" & Me.txtLastKMCounter.Text & "'"
        End If
    End If
    
  If Me.TXTCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.Fullcode like '%" & Me.TXTCode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  ddbo.TblCarsData.Fullcode like '%" & Me.TXTCode.Text & "%'"
        End If
    End If
     If Me.TxtModel.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.Model='" & Me.TxtModel.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.Model='" & Me.TxtModel.Text & "'"
        End If
    End If
    
 If Me.txtBoardNO.Text <> "" Then
        If BolBegine = True Then
'            StrWhere = StrWhere & " AND  dbo.TblCarsData.BoardNO='" & Me.TxtBoardNO.Text & "'"
StrWhere = StrWhere & " AND REPLACE(dbo.TblCarsData.BoardNO, ' ', '')LIKE '%" & Replace(txtBoardNO, " ", "") & "%'"

              
             
        Else
            BolBegine = True
            'StrWhere = " Where  dbo.TblCarsData.BoardNO='" & Me.TxtBoardNO.Text & "'"
            StrWhere = StrWhere & " WHERE REPLACE(dbo.TblCarsData.BoardNO, ' ', '') lIKE '%" & Replace(txtBoardNO, " ", "") & "%'"
        End If
    End If
     If Me.TxtLicenseNO.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.LicenseNO='" & Me.TxtLicenseNO.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.LicenseNO='" & Me.TxtLicenseNO.Text & "'"
        End If
    End If
    
    If val(Me.DcFixedAssets.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarsData.fixedAssetid=" & Me.DcFixedAssets.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarsData.fixedAssetid=" & Me.DcFixedAssets.BoundText & ""
        End If
    End If

    If val(Me.DcEmployee.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.Emp_id=" & Me.DcEmployee.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.Emp_id =" & Me.DcEmployee.BoundText & ""
        End If
    End If
 
 If val(Me.DCGroup.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.CarsTypeId=" & Me.DCGroup.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.CarsTypeId=" & Me.DCGroup.BoundText & ""
        End If
    End If
    If val(Me.DCInsuranceCompanyId.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.InsuranceCompanyId=" & Me.DCInsuranceCompanyId.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.InsuranceCompanyId=" & Me.DCInsuranceCompanyId.BoundText & ""
        End If
    End If
    
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.PurchaseDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.PurchaseDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCarsData.PurchaseDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCarsData.PurchaseDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    
       If SystemOptions.usertype <> UserAdminAll Then
       If BolBegine = True Then
            StrSQL = StrSQL & " AND   TblCarsData.branch_no=" & Current_branch
        Else
            BolBegine = True
            StrSQL = StrSQL & " where   TblCarsData.branch_no=" & Current_branch
        End If
    End If
    
    
    
    
    
    StrSQL = StrSQL & " Order By dbo.TblCarsData.id "
    Set rs = New ADODB.Recordset
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
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
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                .TextMatrix(i, .ColIndex("Capacity")) = IIf(IsNull(rs("Capacity").value), "", rs("Capacity").value)
                 
                If Not (IsNull(rs("PurchaseDate").value)) Then
                    .TextMatrix(i, .ColIndex("PurchaseDate")) = Format(rs("PurchaseDate").value, "yyyy/M/d")
                End If
              If SystemOptions.UserInterface = EnglishInterface Then
               .TextMatrix(i, .ColIndex("FixedAssets_Name")) = IIf(IsNull(rs("FixedAssets_NameE").value), "", rs("FixedAssets_NameE").value)
               .TextMatrix(i, .ColIndex("TBLCarTypes_name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
               .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
              .TextMatrix(i, .ColIndex("Inst_Name")) = IIf(IsNull(rs("Inst_NameE").value), "", rs("Inst_NameE").value)
                Else
                .TextMatrix(i, .ColIndex("FixedAssets_Name")) = IIf(IsNull(rs("FixedAssets_Name").value), "", rs("FixedAssets_Name").value)
                .TextMatrix(i, .ColIndex("TBLCarTypes_name")) = IIf(IsNull(rs("TBLCarTypes_name").value), "", rs("TBLCarTypes_name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Inst_Name")) = IIf(IsNull(rs("Inst_Name").value), "", rs("Inst_Name").value)
              
              End If
              
              .TextMatrix(i, .ColIndex("OperatorN")) = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
              .TextMatrix(i, .ColIndex("FixedassetId")) = IIf(IsNull(rs("FixedassetId").value), 0, rs("FixedassetId").value)
               .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
               .TextMatrix(i, .ColIndex("Model1")) = IIf(IsNull(rs("Model").value), "", rs("Model").value)
                .TextMatrix(i, .ColIndex("VehicleLong")) = IIf(IsNull(rs("VehicleLong").value), "", rs("VehicleLong").value)
               .TextMatrix(i, .ColIndex("EquQty")) = IIf(IsNull(rs("EquQty").value), "", rs("EquQty").value)
               .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
               .TextMatrix(i, .ColIndex("LastKMCounter")) = IIf(IsNull(rs("LastKMCounter").value), "", rs("LastKMCounter").value)
                .TextMatrix(i, .ColIndex("LicenseNO")) = IIf(IsNull(rs("LicenseNO").value), "", rs("LicenseNO").value)
               .TextMatrix(i, .ColIndex("LicenseExpireDateH")) = IIf(IsNull(rs("LicenseExpireDateH").value), "", rs("LicenseExpireDateH").value)
              .TextMatrix(i, .ColIndex("InsuranceExpireDateH")) = IIf(IsNull(rs("InsuranceExpireDateH").value), "", rs("InsuranceExpireDateH").value)
               .TextMatrix(i, .ColIndex("TestExpireDateH")) = IIf(IsNull(rs("TestExpireDateH").value), "", rs("TestExpireDateH").value)
                .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("Notes").value), "", rs("Notes").value)
              .TextMatrix(i, .ColIndex("Emp_mobile")) = IIf(IsNull(rs("Emp_mobile").value), "", rs("Emp_mobile").value)
              .TextMatrix(i, .ColIndex("Emp_Phone")) = IIf(IsNull(rs("Emp_Phone").value), "", rs("Emp_Phone").value)
               
               
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub




Private Sub TxtEquQty_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtEquQty.Text, 1)
End Sub



Private Sub VehicleLong_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.VehicleLong.Text, 1)
End Sub
