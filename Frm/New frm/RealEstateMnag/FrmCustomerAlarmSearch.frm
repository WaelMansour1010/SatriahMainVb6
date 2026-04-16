VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCustomerAlarmSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11505
   Icon            =   "FrmCustomerAlarmSearch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   11505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Caption         =   " «—ÌŒ  «·ÿ·»"
      Height          =   1035
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   3480
      Width           =   3975
      Begin MSComCtl2.DTPicker EndDateTO 
         Height          =   330
         Left            =   1770
         TabIndex        =   44
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker EndDateFrom 
         Height          =   330
         Left            =   1770
         TabIndex        =   43
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal EndDateFromH 
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal EndDateTOH 
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰"
         Height          =   195
         Index           =   6
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   2
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.Frame lbreg 
      Caption         =   " «—ÌŒ  «·ÿ·»"
      Height          =   1035
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3480
      Width           =   3975
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   1770
         TabIndex        =   36
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
         TabIndex        =   37
         Top             =   600
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
         TabIndex        =   38
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToh 
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰"
         Height          =   195
         Index           =   1
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   9720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   5400
      Width           =   825
   End
   Begin VB.TextBox TxtMobile 
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
      Left            =   0
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   5040
      Width           =   5205
   End
   Begin VB.TextBox TxtPeriods 
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
      Left            =   7800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   7440
      Width           =   825
   End
   Begin VB.ComboBox DcbPeriodsID 
      Height          =   315
      ItemData        =   "FrmCustomerAlarmSearch.frx":000C
      Left            =   6960
      List            =   "FrmCustomerAlarmSearch.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7440
      Width           =   735
   End
   Begin VB.TextBox Text15 
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
      Left            =   9720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4680
      Width           =   825
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
      Left            =   9720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5040
      Width           =   825
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCustomerAlarmSearch.frx":0024
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtremark 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Left            =   17040
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8160
      TabIndex        =   4
      Top             =   3600
      Width           =   2235
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":0040
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":03DA
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":0774
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":0B0E
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":0EA8
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":1242
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":15DC
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCustomerAlarmSearch.frx":1B76
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ «‘⁄«—  ”œÌœ/«‰–«—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   2
         Left            =   6135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   5
      Top             =   8160
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   11475
      _cx             =   20241
      _cy             =   4842
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
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCustomerAlarmSearch.frx":1F10
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
      Left            =   1920
      TabIndex        =   10
      Top             =   5520
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   900
      TabIndex        =   11
      Top             =   5520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Index           =   2
      Left            =   -480
      TabIndex        =   13
      Top             =   4410
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
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
      ButtonImage     =   "FrmCustomerAlarmSearch.frx":2207
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DcbIqara 
      Height          =   315
      Left            =   6480
      TabIndex        =   16
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
      Top             =   5040
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcCustomer 
      Height          =   315
      Left            =   6480
      TabIndex        =   19
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
      Top             =   4680
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitNo 
      Height          =   315
      Left            =   60
      TabIndex        =   21
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   4680
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitType 
      Height          =   315
      Left            =   3240
      TabIndex        =   27
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   4680
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitNo1 
      Height          =   315
      Left            =   10440
      TabIndex        =   28
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   9360
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcsupplier 
      Height          =   315
      Left            =   6480
      TabIndex        =   33
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
      Top             =   5400
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbAlarm 
      Height          =   315
      Left            =   8160
      TabIndex        =   49
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   4080
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «·„«·ﬂ"
      Height          =   285
      Index           =   1
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "‰Ê⁄ «·«‘⁄«—"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÊ«· «·„” «Ã—"
      Height          =   195
      Index           =   0
      Left            =   5265
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
      Height          =   405
      Index           =   11
      Left            =   8445
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   7440
      Width           =   1050
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ÊÕœ…"
      Height          =   195
      Index           =   50
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4680
      Width           =   990
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·ÊÕœ…"
      Height          =   195
      Index           =   15
      Left            =   5145
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4680
      Width           =   990
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «·„” √Ã—"
      Height          =   285
      Index           =   5
      Left            =   10605
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄ﬁ«—"
      Height          =   195
      Index           =   4
      Left            =   10305
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   17040
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·«‘⁄«—"
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmCustomerAlarmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer
Public WithEvents FG1 As VSFlex8UCtl.VSFlexGrid
Attribute FG1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.VSFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long


Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
Set rs = New ADODB.Recordset
            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                fg.Clear flexClearScrollable, flexClearEverything
                fg.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            fg.SetFocus

        Case 1
            clear_all Me
            fg.Clear flexClearScrollable, flexClearEverything
            DcbAlarm.Text = ""
            
          
     DtpDateFrom.value = ""
    DtpDateTo.value = ""
  EndDateFrom.value = ""
    EndDateTo.value = ""

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub







Private Sub DcbIqara_Change()
DcbIqara_Click (0)
End Sub

Private Sub DcbIqara_Click(Area As Integer)
   If val(DcbIqara.BoundText) = 0 Then: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
  ' dcsupplier.BoundText = ownerid
 
End Sub




Private Sub dcCustomer_Change()
dcCustomer_Click (0)
End Sub

Private Sub dcCustomer_Click(Area As Integer)
  If val(dcCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcCustomer.BoundText, EmpCode
    Me.Text15.Text = EmpCode
End Sub



Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
  If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub

Private Sub DtpDateFrom_Change()
If DtpDateFrom.value <> "" Then
         DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
         End If

End Sub


Private Sub DtpDateFromH_LostFocus()
 
             VBA.Calendar = vbCalGreg
           Me.DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
       
End Sub

Private Sub DtpDateTo_Change()
If Me.DtpDateTo.value <> "" Then
         DtpDateToH.value = ToHijriDate(DtpDateTo.value)
End If

End Sub


Private Sub DtpDateToH_LostFocus()
  VBA.Calendar = vbCalGreg
           Me.DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub EndDateFrom_Change()
If Me.EndDateFrom.value <> "" Then
         EndDateFromH.value = ToHijriDate(EndDateFrom.value)
End If
End Sub







Private Sub EndDateFromH_LostFocus()
VBA.Calendar = vbCalGreg
           Me.EndDateFrom.value = ToGregorianDate(EndDateFromH.value)
End Sub

Private Sub EndDateTo_Change()
If EndDateTo.value <> "" Then
         EndDateToH.value = ToHijriDate(EndDateTo.value)
         End If

End Sub

Private Sub EndDateTOH_LostFocus()
VBA.Calendar = vbCalGreg
           Me.EndDateTo.value = ToGregorianDate(EndDateToH.value)
End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
       FrmRsCustomerAlarm.Retrive val(fg.TextMatrix(fg.Row, fg.ColIndex("ContNo")))

ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With fg
        
                
                .TextMatrix(Num, .ColIndex("ContNo")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                 .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecDate").value), "", Trim(rs("RecDate").value))
                 .TextMatrix(Num, .ColIndex("RecordDateH")) = IIf(IsNull(rs("RecDateH").value), "", Trim(rs("RecDateH").value))
                 .TextMatrix(Num, .ColIndex("EndDate")) = IIf(IsNull(rs("EndDate").value), "", Trim(rs("EndDate").value))
                 .TextMatrix(Num, .ColIndex("EndDateH")) = IIf(IsNull(rs("EndDateH").value), "", Trim(rs("EndDateH").value))
                
                .TextMatrix(Num, .ColIndex("unitno")) = IIf(IsNull(rs("unitno").value), "", Trim(rs("Nameunitno").value))
            
            .TextMatrix(Num, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", Trim(rs("aqarname").value))
              If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(Num, .ColIndex("Alarm")) = IIf(IsNull(rs("nameAlarm").value), "", Trim(rs("nameAlarm").value))
              .TextMatrix(Num, .ColIndex("OwnerName")) = IIf(IsNull(rs("OwnerName").value), "", Trim(rs("OwnerName").value))
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                     .TextMatrix(Num, .ColIndex("NamUnit")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                .TextMatrix(Num, .ColIndex("Alarm")) = IIf(IsNull(rs("nameAlarmE").value), "", Trim(rs("nameAlarmE").value))
                .TextMatrix(Num, .ColIndex("OwnerName")) = IIf(IsNull(rs("OwnerNamee").value), "", Trim(rs("OwnerNamee").value))
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                     .TextMatrix(Num, .ColIndex("NamUnit")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                   
                .TextMatrix(Num, .ColIndex("Cus_Mobile")) = IIf(IsNull(rs("Cus_Mobile").value), "", Trim(rs("Cus_Mobile").value))



           
                   ' .TextMatrix(Num, .ColIndex("AccountNo")) = IIf(IsNull(rs("AccountNo").value), "", Trim(rs("AccountNo").value))
               
               '    .TextMatrix(Num, .ColIndex("Water")) = IIf(IsNull(rs("Water").value), "", Trim(rs("Water").value))
           
               '    .TextMatrix(Num, .ColIndex("Electricity")) = IIf(IsNull(rs("Electricity").value), "", Trim(rs("Electricity").value))
          'TextMatrix(Num, .ColIndex("Phone")) = IIf(IsNull(rs("Phone").value), "", Trim(rs("Phone").value))
          

                

              '.TextMatrix(Num, .ColIndex("TotalContract")) = IIf(IsNull(rs("TotalContract").value), "", Trim(rs("TotalContract").value))
              '    .TextMatrix(Num, .ColIndex("PayAmini")) = IIf(IsNull(rs("PayAmini").value), "", Trim(rs("PayAmini").value))
           '
           '      .TextMatrix(Num, .ColIndex("CommiValue")) = IIf(IsNull(rs("CommiValue").value), "", (rs("CommiValue").value))
           ''       .TextMatrix(Num, .ColIndex("InsuranceValue")) = IIf(IsNull(rs("InsuranceValue").value), "", Trim(rs("InsuranceValue").value))
            
            End With

            rs.MoveNext
        Next Num

         fg.AutoSize 0, fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub









Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
   Dcombos.GetAlarm Me.DcbAlarm
   Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
   Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
   Dcombos.GetIqar DcbIqara
   Dcombos.getAkarUnit Me.DcbUnitType
   Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
          
          DcbAlarm.Text = ""
     DtpDateFrom.value = ""
    DtpDateTo.value = ""
  EndDateFrom.value = ""
    EndDateTo.value = ""

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
   ' My_SQL = "select UserID,UserName From tblUsers "
   ' fill_combo DCUser, My_SQL
 
  'Dcombos.GetIqarUnit 1, DcbUnitNo
  'combos.GetBranches dcBranch

 'Dcombos.GetSalesRepData Me.DcboEmp   Set cSearch = New clsDCboSearch
   ' My_SQL = " select CountryID,CountryName from TblCountriesData"
 
   ' fill_combo Me.DataCombo4, My_SQL
   ' RetrunType = -1
 
    CenterForm Me

    FormPostion Me, GetPostion
    fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    DBCboClientName.BoundText = ""
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

 StrSQL = " SELECT     dbo.TblCustomerAlarm.RecDateH, dbo.TblCustomerAlarm.RecDate, dbo.TblCustomerAlarm.ContNo, dbo.TblCustomerAlarm.UserID, dbo.TblCustomerAlarm.Cus_Mobile,"
   StrSQL = StrSQL & "                      dbo.TblCustomerAlarm.Des, dbo.TblCustomerAlarm.EndDateH, dbo.TblCustomerAlarm.EndDate, dbo.TblCustomerAlarm.CusID, TblCustemers_1.CusName,"
   StrSQL = StrSQL & "                     TblCustemers_1.CusNamee, dbo.TblCustomerAlarm.ownerid, TblCustemers_1.CusName AS OwnerName, TblCustemers_1.CusNamee AS OwnerNameE,"
   StrSQL = StrSQL & "                     dbo.TblCustomerAlarm.AqrID, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblCustomerAlarm.UnitNo, dbo.TblAqarDetai.unitno AS Nameunitno,"
  StrSQL = StrSQL & "                      dbo.TblCustomerAlarm.ID, dbo.TblCustomerAlarm.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblCustomerAlarm.Alarm,"
  StrSQL = StrSQL & "                      dbo.TblAlarmType.name AS nameAlarm, dbo.TblAlarmType.namee AS nameAlarmE"
 StrSQL = StrSQL & "  FROM         dbo.TblCustomerAlarm LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblAlarmType ON dbo.TblCustomerAlarm.Alarm = dbo.TblAlarmType.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblCustomerAlarm.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblCustomerAlarm.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblCustomerAlarm.AqrID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblCustomerAlarm.ownerid = TblCustemers_1.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_2 ON dbo.TblCustomerAlarm.CusID = TblCustemers_2.CusID"
 StrSQL = StrSQL & " "
    StrSQL = StrSQL & "   where 1=1 "
  If Me.TXTOrDer_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.ID ='" & Me.TXTOrDer_no.Text & "'"
 
    End If
    
    If val(DcbUnitType.BoundText) <> 0 And DcbUnitType.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.UnitType =" & DcbUnitType.BoundText & ""
 
    End If
   
   If val(Me.DcbUnitNo.BoundText) <> 0 And Me.DcbUnitNo.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.UnitNo =" & val(Me.DcbUnitNo.BoundText)
        End If

    If Me.DcbIqara.BoundText <> "" And Me.DcbIqara.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.Aqarid =" & val(Me.DcbIqara.BoundText)
 
    End If
    If Me.dcCustomer.BoundText <> "" And Me.dcCustomer.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.CusID =" & val(Me.dcCustomer.BoundText)
 
    End If
  If Me.dcsupplier.BoundText <> "" And Me.dcsupplier.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.ownerid =" & val(Me.dcsupplier.BoundText)
 
    End If
      If val(Me.DcbAlarm.BoundText) <> 0 And Me.DcbAlarm.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.Alarm =" & val(Me.DcbAlarm.BoundText)
 
    End If
    
  
If Me.TxtMobile.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustomerAlarm.Cus_Mobile ='" & Me.TxtMobile.Text & "'"
 
    End If
       If Not IsNull(Me.DtpDateFrom.value) Then
    
            StrWhere = StrWhere & " AND  dbo.TblCustomerAlarm.RecDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
   End If
       If Not IsNull(Me.DtpDateTo.value) Then
    
            StrWhere = StrWhere & " AND  dbo.TblCustomerAlarm.RecDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
   End If
   
 
         If Not IsNull(Me.EndDateFrom.value) Then
    
            StrWhere = StrWhere & " AND  dbo.TblCustomerAlarm.EndDate >=" & SQLDate(Me.EndDateFrom.value, True) & ""
   End If
       If Not IsNull(Me.EndDateTo.value) Then
    
            StrWhere = StrWhere & " AND  dbo.TblCustomerAlarm.EndDate <=" & SQLDate(Me.EndDateTo.value, True) & ""
   End If

    StrWhere = StrWhere + " order by dbo.TblCustomerAlarm.ID"

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is fg Then
            If Not fg.TextMatrix(fg.Row, 1) = "" Then
                Fg_Click
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub







Private Sub ChangeLang()
    Me.Caption = "Search For IqarFilter waiver"
    Label1(2).Caption = Me.Caption
    lbl(0).Caption = "Filte No"
 lbl(15).Caption = "Type"
  lbl(2).Caption = "Op Date"
  lbl(50).Caption = "Unite"
  lbl(19).Caption = "Insurance"
    lbl(4).Caption = "Iqar"
        lbl(21).Caption = "Electricity"
        lbl(28).Caption = "No Day"
           lbl(20).Caption = "Account No"
         lbl(5).Caption = "Renter "
         lbl(22).Caption = "End Rent "
          lbl(23).Caption = "DateFiltering  "
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.fg
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
         .TextMatrix(0, .ColIndex("ContNo")) = "Filte No  "
        .TextMatrix(0, .ColIndex("RecordDate")) = "OpDateAD"
        .TextMatrix(0, .ColIndex("RecordDateH")) = " OpDateHj"
          .TextMatrix(0, .ColIndex("EndDate")) = "EndDateAD"
           .TextMatrix(0, .ColIndex("EndDateH")) = "EndDateHj"
            .TextMatrix(0, .ColIndex("FilterDate")) = "FilterDateAD"
           .TextMatrix(0, .ColIndex("FilterDateH")) = "FilterDateHj"
           
            .TextMatrix(0, .ColIndex("aqarname")) = " Iqarname"
          .TextMatrix(0, .ColIndex("NamUnit")) = "UnitType"
           .TextMatrix(0, .ColIndex("unitno")) = "UnitNo"
            .TextMatrix(0, .ColIndex("CusName")) = "Renter"
           .TextMatrix(0, .ColIndex("Insurance")) = "Insurance"
       .TextMatrix(0, .ColIndex("BillPrice")) = " Electricity"
          .TextMatrix(0, .ColIndex("DayNo")) = "DayNo"
           .TextMatrix(0, .ColIndex("AccountNo")) = "AccountNo"
           
  
          .AutoSize 0, .Cols - 1, False
    End With

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text1.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
    End If
End Sub



Private Sub Text15_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
    End If
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


