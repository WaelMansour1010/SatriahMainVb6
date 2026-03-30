VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmIqarWaiverSet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16065
   Icon            =   "FrmIqarWaiverSet.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   16065
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
      TabIndex        =   12
      Top             =   7440
      Width           =   825
   End
   Begin VB.ComboBox DcbPeriodsID 
      Height          =   315
      ItemData        =   "FrmIqarWaiverSet.frx":000C
      Left            =   6960
      List            =   "FrmIqarWaiverSet.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   7440
      Width           =   735
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   6
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
      MICON           =   "FrmIqarWaiverSet.frx":0024
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
      TabIndex        =   5
      Top             =   4680
      Width           =   7830
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   16065
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
               Picture         =   "FrmIqarWaiverSet.frx":0040
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":03DA
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":0774
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":0B0E
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":0EA8
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":1242
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":15DC
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarWaiverSet.frx":1B76
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   480
         Picture         =   "FrmIqarWaiverSet.frx":1F10
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ  ’›Ì… Ê ‰«“· ⁄‰ «·⁄ﬁÊœ"
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
         Left            =   10575
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   3
      Top             =   8160
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   5280
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
      TabIndex        =   8
      Top             =   5280
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
      TabIndex        =   9
      Top             =   5280
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
      TabIndex        =   10
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
      ButtonImage     =   "FrmIqarWaiverSet.frx":5B78
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DcbUnitNo1 
      Height          =   315
      Left            =   10440
      TabIndex        =   14
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   9360
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   16095
      Begin VB.TextBox txtorder_no 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   13800
         TabIndex        =   23
         Top             =   2880
         Width           =   1035
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
         Left            =   13800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3360
         Width           =   1035
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
         Left            =   13800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3840
         Width           =   1065
      End
      Begin VB.TextBox TxtInsuranceValue 
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
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2880
         Width           =   1965
      End
      Begin VB.TextBox TxtLateDate 
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
         Left            =   3240
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2880
         Width           =   1965
      End
      Begin VB.TextBox TxtElectricity 
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
         Left            =   6480
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2880
         Width           =   1965
      End
      Begin VB.TextBox TxtAccountNo 
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
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3360
         Width           =   1965
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2745
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   16035
         _cx             =   28284
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
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmIqarWaiverSet.frx":6112
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
      Begin MSComCtl2.DTPicker XPDtbBill 
         Height          =   315
         Left            =   11040
         TabIndex        =   24
         Top             =   2880
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64421889
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcbIqara 
         Height          =   315
         Left            =   9720
         TabIndex        =   25
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   3360
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcCustomer 
         Height          =   315
         Left            =   9720
         TabIndex        =   26
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   3840
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitNo 
         Height          =   315
         Left            =   3240
         TabIndex        =   27
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3360
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker EndDate 
         Height          =   315
         Left            =   6720
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   64421891
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker FilterDate 
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   64421891
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo DcbUnitType 
         Height          =   315
         Left            =   6480
         TabIndex        =   30
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3360
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
         Height          =   315
         Left            =   9720
         TabIndex        =   31
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal EndDateH 
         Height          =   315
         Left            =   5160
         TabIndex        =   32
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal FilterDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «· ’›ÌÂ"
         Height          =   375
         Index           =   0
         Left            =   14640
         TabIndex        =   46
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ «·⁄„·ÌÂ"
         Height          =   375
         Index           =   2
         Left            =   12720
         TabIndex        =   45
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblitemid 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   195
         Index           =   4
         Left            =   14865
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„” √Ã—"
         Height          =   285
         Index           =   5
         Left            =   15165
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   15
         Left            =   8385
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Â«Ì… «·«ÌÃ«—"
         Height          =   285
         Index           =   22
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· ’›ÌÂ"
         Height          =   405
         Index           =   23
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   3840
         Width           =   1050
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÊÕœ…"
         Height          =   195
         Index           =   50
         Left            =   5340
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " √„Ì‰"
         Height          =   195
         Index           =   19
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «Ì«„ «· «ŒÌ—"
         Height          =   195
         Index           =   28
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÂ—»«¡"
         Height          =   195
         Index           =   21
         Left            =   8385
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·Õ”«»"
         Height          =   195
         Index           =   20
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   3360
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   720
      Width           =   16095
      Begin VB.ComboBox DcbTypID2 
         Height          =   315
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   4080
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.ComboBox DcbTypID 
         Height          =   315
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   2880
         Width           =   3345
      End
      Begin VB.TextBox Text3 
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
         Left            =   13800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   3840
         Width           =   1065
      End
      Begin VB.TextBox Text2 
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
         Left            =   13800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   3360
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   13800
         TabIndex        =   48
         Top             =   2880
         Width           =   1035
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2745
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   16035
         _cx             =   28284
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
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmIqarWaiverSet.frx":6491
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
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   315
         Left            =   11040
         TabIndex        =   52
         Top             =   2880
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64421889
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcbIqara2 
         Height          =   315
         Left            =   9720
         TabIndex        =   53
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   3360
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcCustomer2 
         Height          =   315
         Left            =   9720
         TabIndex        =   54
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   3840
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   315
         Left            =   6600
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   64421891
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker ToFiter 
         Height          =   315
         Left            =   1980
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   64421891
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo DcbUnitType2 
         Height          =   315
         Left            =   4980
         TabIndex        =   57
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3360
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Dynamic_Byte.NourHijriCal FromDateH 
         Height          =   315
         Left            =   9720
         TabIndex        =   58
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal ToDateH 
         Height          =   315
         Left            =   4980
         TabIndex        =   59
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal ToFiterH 
         Height          =   315
         Left            =   360
         TabIndex        =   60
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
      End
      Begin MSDataListLib.DataCombo DcbUnitNo2 
         Height          =   315
         Left            =   360
         TabIndex        =   71
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   3360
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker FromFiter 
         Height          =   315
         Left            =   6600
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   64421891
         CurrentDate     =   37140
      End
      Begin Dynamic_Byte.NourHijriCal FromFiterH 
         Height          =   315
         Left            =   4980
         TabIndex        =   74
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· ’›Ì… „‰"
         Height          =   405
         Index           =   13
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   3840
         Width           =   1290
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "‰Ê⁄ «·⁄„·Ì…"
         Height          =   195
         Index           =   14
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÊÕœ…"
         Height          =   195
         Index           =   12
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï  «—ÌŒ "
         Height          =   195
         Index           =   10
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï  «—ÌŒ"
         Height          =   285
         Index           =   9
         Left            =   8205
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2880
         Width           =   1170
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   8
         Left            =   8385
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„” √Ã—"
         Height          =   285
         Index           =   7
         Left            =   15045
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   195
         Index           =   6
         Left            =   14865
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰  «—ÌŒ"
         Height          =   375
         Index           =   3
         Left            =   12720
         TabIndex        =   62
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «· ’›ÌÂ"
         Height          =   375
         Index           =   1
         Left            =   14640
         TabIndex        =   61
         Top             =   2880
         Width           =   1215
      End
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
      TabIndex        =   13
      Top             =   7440
      Width           =   1050
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   17040
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   8280
      Width           =   735
   End
End
Attribute VB_Name = "FrmIqarWaiverSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Indxx As Integer
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch
Private m_DcboItems As DataCombo
Public m_RetrunType As Integer
Public WithEvents Fg1 As VSFlex8UCtl.VSFlexGrid
Attribute Fg1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.VSFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long


Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If
If m_RetrunType = 8 Or m_RetrunType = 808 Then
     rs.Open Build_Sql2, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If rs.RecordCount < 1 Then
                VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                VSFlexGrid1.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive2
            VSFlexGrid1.SetFocus
Else

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2

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
            FG.SetFocus
End If
        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
     XPDtbBill.value = ""
    EndDate.value = ""
FilterDate.value = ""
fromDate.value = ""
toDate.value = ""
FromFiter.value = ""
ToFiter.value = ""
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub
Private Sub DcbIqara_Click(Area As Integer)
   If val(DcbIqara.BoundText) = 0 Then: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    Me.TxtSearch.Text = EmpCode
End Sub
Private Sub DcbUnitType2_Change()
DcbUnitType2_Click (0)
End Sub

Private Sub DcbUnitType2_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
Set Dcombos = New ClsDataCombos
If val(DcbIqara2.BoundText) > 0 Then
idd = val(DcbIqara2.BoundText)
idd1 = val(DcbUnitType2.BoundText)
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo2, "R"
End If
End Sub

Private Sub dcCustomer_Click(Area As Integer)
  If val(dcCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcCustomer.BoundText, EmpCode
    Me.Text15.Text = EmpCode
End Sub



Private Sub dcCustomer2_Change()
  If val(dcCustomer2.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , dcCustomer2.BoundText, EmpCode
    Me.Text3.Text = EmpCode
End Sub

Private Sub dcCustomer2_Click(Area As Integer)
dcCustomer2_Change
End Sub

Private Sub ENDDATE_Change()
If Not IsNull(EndDate.value) Then
         EndDateH.value = ToHijriDate(EndDate.value)
End If
End Sub

Private Sub ENDDATEH_LostFocus()
             VBA.Calendar = vbCalGreg
           Me.EndDate.value = ToGregorianDate(EndDateH.value)
End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
   If m_RetrunType = 5 Then
            
            
            If val(FG.TextMatrix(FG.Row, FG.ColIndex("net"))) > 0 Then
                    FrmCashing1.DcbIqara.BoundText = Abs(val(FG.TextMatrix(FG.Row, FG.ColIndex("iqarid"))))
                    FrmCashing1.TxtFilterNo.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
                    
                    FrmCashing1.TXtFilter.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("net")))
                    FrmCashing1.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("CusID")))
                    FrmCashing1.DcbUnitNo.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
                    FrmCashing1.Dtaefilter.value = FG.TextMatrix(FG.Row, FG.ColIndex("RecordDate"))
                    FrmCashing1.XPTxtVal.Text = Abs(val(FG.TextMatrix(FG.Row, FG.ColIndex("net"))))
                    FrmCashing1.txtTotalinsuranceS.Text = Abs(val(FG.TextMatrix(FG.Row, FG.ColIndex("TotalinsuranceS"))))
                    FrmCashing1.GetWonerID (val(FG.TextMatrix(FG.Row, FG.ColIndex("iqarid"))))
                    
                    
            Else
               MsgBox "Â–… «· ’›Ì… ·«ÌÊÃœ ⁄·ÌÂ« „»«·€ ··ﬁ»÷ "
            End If
    
    
     ElseIf m_RetrunType = 7 Then
               If val(FG.TextMatrix(FG.Row, FG.ColIndex("net"))) < 0 Then
               
                FrmPayments2.TxtFilterNo.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
                FrmPayments2.TXtFilter.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("net")))
                FrmPayments2.GetWonerID (val(FG.TextMatrix(FG.Row, FG.ColIndex("iqarid"))))
                FrmPayments2.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("CusID")))
                ''//
                FrmPayments2.DcbIqara.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("iqarid")))
                FrmPayments2.DcbUnitType.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("unitid")))
                FrmPayments2.DcbUnitNo.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("unitnoid")))
                
                FrmPayments2.XPTxtVal.Text = Abs(val(FG.TextMatrix(FG.Row, FG.ColIndex("net"))))
            
                Else
                   MsgBox "Â–… «· ’›Ì… ·«ÌÊÃœ ⁄·ÌÂ« „»«·€ ··œ›⁄"
                End If
    
    
    ''//
    
    Else
       FrmWaiverSettlement.Retrive val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
End If
ErrTrap:
End Sub
Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                 .TextMatrix(Num, .ColIndex("iqarid")) = IIf(IsNull(rs("Aqarid").value), "", rs("Aqarid").value)
                 .TextMatrix(Num, .ColIndex("unitid")) = IIf(IsNull(rs("UnitType").value), "", rs("UnitType").value)
                 .TextMatrix(Num, .ColIndex("unitnoid")) = IIf(IsNull(rs("IDUNo").value), "", rs("IDUNo").value)
                 .TextMatrix(Num, .ColIndex("TotalinsuranceS")) = IIf(IsNull(rs("TotalinsuranceS").value), "", rs("TotalinsuranceS").value)
                 .TextMatrix(Num, .ColIndex("CusID")) = IIf(IsNull(rs("RenterID").value), "", rs("RenterID").value)
                 .TextMatrix(Num, .ColIndex("Aqarid")) = IIf(IsNull(rs("IDUNo").value), "", rs("IDUNo").value)
                 .TextMatrix(Num, .ColIndex("net")) = IIf(IsNull(rs("net").value), "", rs("net").value)
                 .TextMatrix(Num, .ColIndex("ContNo")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                 .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", Trim(rs("RecordDate").value))
                 .TextMatrix(Num, .ColIndex("RecordDateH")) = IIf(IsNull(rs("RecordDateH").value), "", Trim(rs("RecordDateH").value))
                 .TextMatrix(Num, .ColIndex("EndDate")) = IIf(IsNull(rs("EndDate").value), "", Trim(rs("EndDate").value))
                 .TextMatrix(Num, .ColIndex("EndDateH")) = IIf(IsNull(rs("EndDateH").value), "", Trim(rs("EndDateH").value))
                 .TextMatrix(Num, .ColIndex("FilterDate")) = IIf(IsNull(rs("FilterDate").value), "", Trim(rs("FilterDate").value))
                 .TextMatrix(Num, .ColIndex("FilterDateH")) = IIf(IsNull(rs("FilterDateH").value), "", Trim(rs("FilterDateH").value))
                 .TextMatrix(Num, .ColIndex("unitno")) = IIf(IsNull(rs("unitno").value), "", Trim(rs("unitno").value))
                 .TextMatrix(Num, .ColIndex("Insurance")) = IIf(IsNull(rs("Insurance").value), "", Trim(rs("Insurance").value))
                 .TextMatrix(Num, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", Trim(rs("aqarname").value))
              If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                     .TextMatrix(Num, .ColIndex("NamUnit")) = IIf(IsNull(rs("NamUnit").value), "", rs("NamUnit").value)
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                     .TextMatrix(Num, .ColIndex("NamUnit")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                   
                .TextMatrix(Num, .ColIndex("BillPrice")) = IIf(IsNull(rs("BillPrice").value), "", Trim(rs("BillPrice").value))

             
.TextMatrix(Num, .ColIndex("DayNo")) = IIf(IsNull(rs("DayNo").value), "", Trim(rs("DayNo").value))
           
                    .TextMatrix(Num, .ColIndex("AccountNo")) = IIf(IsNull(rs("AccountNo").value), "", Trim(rs("AccountNo").value))
               
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

         FG.AutoSize 0, FG.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub Retrive2()
    Dim Num As Integer
    On Error GoTo ErrTrap
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        VSFlexGrid1.Rows = rs.RecordCount + 1
With VSFlexGrid1
        For Num = 1 To rs.RecordCount
            .TextMatrix(Num, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
            .TextMatrix(Num, .ColIndex("iqarid")) = IIf(IsNull(rs("AqarID").value), "", rs("AqarID").value)
            .TextMatrix(Num, .ColIndex("unitid")) = IIf(IsNull(rs("UnitTypID").value), "", rs("UnitTypID").value)
            .TextMatrix(Num, .ColIndex("unitnoid")) = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
            .TextMatrix(Num, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
            .TextMatrix(Num, .ColIndex("Aqarid")) = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
            .TextMatrix(Num, .ColIndex("ContNo")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
            .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", Trim(rs("RecordDate").value))
            .TextMatrix(Num, .ColIndex("RecordDateH")) = IIf(IsNull(rs("RecordDateH").value), "", Trim(rs("RecordDateH").value))
            .TextMatrix(Num, .ColIndex("FilterDate")) = IIf(IsNull(rs("PayedDate").value), "", Trim(rs("PayedDate").value))
            .TextMatrix(Num, .ColIndex("FilterDateH")) = IIf(IsNull(rs("FromDateH").value), "", Trim(rs("FromDateH").value))
            .TextMatrix(Num, .ColIndex("unitno")) = IIf(IsNull(rs("unitno").value), "", Trim(rs("unitno").value))
            .TextMatrix(Num, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", Trim(rs("aqarname").value))
            DcbTypID.ListIndex = IIf(IsNull(rs("TypID").value), -1, Trim(rs("TypID").value))
            .TextMatrix(Num, .ColIndex("TypeName")) = DcbTypID.Text
              If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                    .TextMatrix(Num, .ColIndex("NamUnit")) = IIf(IsNull(rs("Unitname").value), "", rs("Unitname").value)
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                    .TextMatrix(Num, .ColIndex("NamUnit")) = IIf(IsNull(rs("UnitnameE").value), "", rs("UnitnameE").value)
               End If
            rs.MoveNext
        Next Num

        .AutoSize 0, FG.Cols - 1, False
       End With
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub



Private Sub FilterDate_Change()
If Not IsNull(FilterDate.value) Then
         FilterDateH.value = ToHijriDate(FilterDate.value)
End If
End Sub
Private Sub FilterDateH_LostFocus()
             VBA.Calendar = vbCalGreg
           Me.FilterDateH.value = ToGregorianDate(FilterDateH.value)
End Sub
Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       Frame1.Visible = False
       Frame2.Visible = False
If m_RetrunType = 8 Or m_RetrunType = 808 Then
Frame2.Visible = True
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer2
    Dcombos.GetIqar DcbIqara2
    Dcombos.getAkarUnit Me.DcbUnitType2
    fromDate.value = Date
    toDate.value = Date
    FromFiter.value = Date
    ToFiter.value = Date
If SystemOptions.UserInterface = ArabicInterface Then
With DcbTypID
.Clear
.AddItem " ’›Ì…"
.AddItem "›« Ê—… ﬂÂ—»«¡"
End With
With DcbTypID2
.Clear
.AddItem " ’›Ì…"
.AddItem "›« Ê—… ﬂÂ—»«¡"
End With
Else
With DcbTypID
.Clear
.AddItem "Evacuation"
.AddItem "Electricity"
End With
With DcbTypID2
.Clear
.AddItem "Evacuation"
.AddItem "Electricity"
End With
End If
Else
  Frame1.Visible = True
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
End If
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    XPDtbBill.value = ""
    EndDate.value = ""
FilterDate.value = ""

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
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
Private Function Build_Sql2()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

 StrSQL = " SELECT     dbo.TblOtheExpensAqar.ID, dbo.TblOtheExpensAqar.RecordDateH, dbo.TblOtheExpensAqar.RecordDate, dbo.TblOtheExpensAqar.Valuee, "
 StrSQL = StrSQL & "                      dbo.TblOtheExpensAqar.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblOtheExpensAqar.BillNo,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.AccountNo, dbo.TblOtheExpensAqar.AccountBank, dbo.TblOtheExpensAqar.Remarks, dbo.TblOtheExpensAqar.PayedDateH,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.PayedDate, dbo.TblOtheExpensAqar.FromDateH, dbo.TblOtheExpensAqar.FromDate, dbo.TblOtheExpensAqar.ToDateH,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.ToDate, dbo.TblOtheExpensAqar.TypID, dbo.TblOtheExpensAqar.Mobile, dbo.TblOtheExpensAqar.StatusOper,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.RemainRent, dbo.TblOtheExpensAqar.Electricity, dbo.TblOtheExpensAqar.Maintenance, dbo.TblOtheExpensAqar.MaintCondition,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.MaintDoors, dbo.TblOtheExpensAqar.MaintKitchen, dbo.TblOtheExpensAqar.MaintClean, dbo.TblOtheExpensAqar.MaintOther,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.Insurance, dbo.TblOtheExpensAqar.DelayDay, dbo.TblOtheExpensAqar.Noliquidation, dbo.TblOtheExpensAqar.Paints,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.Windows, dbo.TblOtheExpensAqar.Total, dbo.TblOtheExpensAqar.Net, dbo.TblOtheExpensAqar.Name, dbo.TblOtheExpensAqar.TotalAfterIns,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.Discount, dbo.TblOtheExpensAqar.NoteSerial1, dbo.TblOtheExpensAqar.Prefix, dbo.TblOtheExpensAqar.NoteSerial,"
 StrSQL = StrSQL & "                     dbo.TblOtheExpensAqar.AqarID, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblOtheExpensAqar.UnitTypID, dbo.TblAkarUnit.name AS Unitname,"
 StrSQL = StrSQL & "                     dbo.TblAkarUnit.namee AS UnitnameE, dbo.TblOtheExpensAqar.UnitID, dbo.TblAqarDetai.unitno, dbo.TblOtheExpensAqar.CusID, dbo.TblCustemers.CusName,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblOtheExpensAqar.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode,"
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee"
 StrSQL = StrSQL & " FROM         dbo.TblOtheExpensAqar LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblOtheExpensAqar.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.TblOtheExpensAqar.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqarDetai ON dbo.TblOtheExpensAqar.UnitID = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAkarUnit ON dbo.TblOtheExpensAqar.UnitTypID = dbo.TblAkarUnit.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqar ON dbo.TblOtheExpensAqar.AqarID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblOtheExpensAqar.BranchID = dbo.TblBranchesData.branch_id"
 StrSQL = StrSQL & "   where 1=1 "
  If Me.Text1.Text <> "" Then
        StrWhere = StrWhere + " and dbo.TblOtheExpensAqar.NoteSerial1 ='" & Me.Text1.Text & "'"
    End If
    If val(DcbTypID.ListIndex) <> -1 And DcbTypID.Text <> "" Then
        StrWhere = StrWhere + " and dbo.TblOtheExpensAqar.TypID =" & val(DcbTypID.ListIndex) & ""
    End If
       If val(DcbUnitType2.BoundText) <> 0 And DcbUnitType2.Text <> "" Then
        StrWhere = StrWhere + " and dbo.TblOtheExpensAqar.UnitTypID =" & val(DcbUnitType2.BoundText) & ""
    End If
    
   If val(Me.DcbUnitNo2.BoundText) <> 0 And Me.DcbUnitNo2.Text <> "" Then
        StrWhere = StrWhere + " and dbo.TblOtheExpensAqar.UnitID =" & val(Me.DcbUnitNo2.BoundText)
    End If

    If Me.DcbIqara2.BoundText <> "" And Me.DcbIqara2.Text <> "" Then
        StrWhere = StrWhere + " and dbo.TblOtheExpensAqar.AqarID =" & val(Me.DcbIqara2.BoundText)
    End If
    If Me.dcCustomer2.BoundText <> "" And Me.dcCustomer2.Text <> "" Then
        StrWhere = StrWhere + " and dbo.TblOtheExpensAqar.CusID =" & val(Me.dcCustomer2.BoundText)
    End If
 
   If Not IsNull(Me.fromDate.value) Then
            StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.RecordDate >=" & SQLDate(Me.fromDate.value, True) & ""
   End If
   If Not IsNull(Me.toDate.value) Then
            StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.RecordDate <=" & SQLDate(Me.toDate.value, True) & ""
   End If

   If Not IsNull(Me.FromFiter.value) Then
            StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.PayedDate >=" & SQLDate(Me.FromFiter.value, True) & ""
   End If
   If Not IsNull(Me.ToFiter.value) Then
            StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.PayedDate <=" & SQLDate(Me.ToFiter.value, True) & ""
   End If
   
    StrWhere = StrWhere + " order by dbo.TblOtheExpensAqar.ID"

    Build_Sql2 = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function
Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

 StrSQL = "SELECT     dbo.TblFiterWaiver.TotalinsuranceS ,dbo.TblFiterWaiver.ID, dbo.TblFiterWaiver.RecordDateH, dbo.TblFiterWaiver.RecordDate, dbo.TblFiterWaiver.BranchID, dbo.TblBranchesData.branch_name, "
         StrSQL = StrSQL & "                dbo.TblBranchesData.branch_namee, dbo.TblFiterWaiver.BulidID, dbo.TblFiterWaiver.RenterID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
        StrSQL = StrSQL & "                 dbo.TblFiterWaiver.Insurance, dbo.TblFiterWaiver.EndDateH, dbo.TblFiterWaiver.EndDate, dbo.TblFiterWaiver.FilterDateH, dbo.TblFiterWaiver.FilterDate,"
       StrSQL = StrSQL & "                  dbo.TblFiterWaiver.BillPrice, dbo.TblFiterWaiver.AccountNo, dbo.TblFiterWaiver.DayNo, dbo.TblFiterWaiver.AmountDely, dbo.TblFiterWaiver.OFRenter,"
       StrSQL = StrSQL & "                  dbo.TblFiterWaiver.ForRenter, dbo.TblAqarDetai.Id AS IDUNo, dbo.TblFiterWaiver.ApartmentID, dbo.TblAqarDetai.unitno, dbo.TblFiterWaiver.unittype,"
     StrSQL = StrSQL & "                    dbo.TblAkarUnit.name AS NamUnit, dbo.TblAkarUnit.namee, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname, dbo.TblFiterWaiver.net"
  StrSQL = StrSQL & "  FROM         dbo.TblCustemers RIGHT OUTER JOIN"
      StrSQL = StrSQL & "                   dbo.TblAqarDetai RIGHT OUTER JOIN"
      StrSQL = StrSQL & "                   dbo.TblFiterWaiver LEFT OUTER JOIN"
     StrSQL = StrSQL & "                    dbo.TblAkarUnit ON dbo.TblFiterWaiver.unittype = dbo.TblAkarUnit.id ON dbo.TblAqarDetai.Id = dbo.TblFiterWaiver.ApartmentID ON"
     StrSQL = StrSQL & "                    dbo.TblCustemers.CusID = dbo.TblFiterWaiver.RenterID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                    dbo.TblAqar ON dbo.TblFiterWaiver.BulidID = dbo.TblAqar.Aqarid RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblFiterWaiver.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " Where (dbo.TblFiterWaiver.id Is Not Null)"
    'StrSQL = StrSQL & "   where 1=1 "
  If Me.TXTOrDer_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.ID ='" & Me.TXTOrDer_no.Text & "'"
 
    End If
    
    If val(DcbUnitType.BoundText) <> 0 And DcbUnitType.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.UnitType =" & DcbUnitType.BoundText & ""
 
    End If
   
   If val(Me.DcbUnitNo.BoundText) <> 0 And Me.DcbUnitNo.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqarDetai.Id =" & val(Me.DcbUnitNo.BoundText)
 
    End If
           If Me.TxtInsuranceValue.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.Insurance ='" & Me.TxtInsuranceValue.Text & "'"
 
    End If
    If Me.DcbIqara.BoundText <> "" And Me.DcbIqara.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.Aqarid =" & val(Me.DcbIqara.BoundText)
 
    End If
    If Me.dcCustomer.BoundText <> "" And Me.dcCustomer.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.RenterID =" & val(Me.dcCustomer.BoundText)
 
    End If
                 If Me.TxtElectricity.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.Electricity ='" & Me.TxtElectricity.Text & "'"
 
    End If
    
  
If Me.TxtLateDate.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.DayNo ='" & Me.TxtLateDate.Text & "'"
 
    End If
      If Me.TxtAccountNo.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblFiterWaiver.AccountNo ='" & Me.TxtAccountNo.Text & "'"
 
    End If
 
         If Not IsNull(Me.EndDate.value) Then
    
            StrWhere = StrWhere & " AND dbo.TblFiterWaiver.EndDate <=" & SQLDate(Me.EndDate.value, True) & ""
   End If
       If Not IsNull(Me.FilterDate.value) Then
    
            StrWhere = StrWhere & " AND dbo.TblFiterWaiver.FilterDate >=" & SQLDate(Me.FilterDate.value, True) & ""
   End If

    StrWhere = StrWhere + " order by dbo.TblFiterWaiver.ID"

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.Row, 1) = "" Then
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
    With Me.FG
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



Private Sub FromDate_Change()
If Not IsNull(fromDate.value) Then
FromdateH.value = ToHijriDate(fromDate.value)
End If
End Sub

Private Sub Fromdateh_LostFocus()
    VBA.Calendar = vbCalGreg
            fromDate.value = ToGregorianDate(FromdateH.value)
End Sub

Private Sub FromFiter_Change()
If Not IsNull(FromFiter.value) Then
FromFiterH.value = ToHijriDate(FromFiter.value)
End If
End Sub

Private Sub FromFiterH_LostFocus()
   VBA.Calendar = vbCalGreg
            FromFiter.value = ToGregorianDate(FromFiterH.value)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Double

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text1.Text, EmpID
        dcsupplier.BoundText = EmpID
    End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub NourHijriCal1_LostFocus()
  '  If Me.TxtModFlg.text <> "R" Then
             VBA.Calendar = vbCalGreg
           Me.XPDtbBill.value = ToGregorianDate(NourHijriCal1.value)
  '         End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
Dim EmpID As Double

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  Dim AqrID As Double
    If KeyAscii = vbKeyReturn Then
        GetIqarCode Text2.Text, AqrID
        DcbIqara2.BoundText = AqrID
        'DcbIqara2_Click (0)
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 Dim EmpID As Double
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text3.Text, EmpID, , , 56
        dcCustomer2.BoundText = EmpID
    End If
End Sub

Private Sub ToDate_Change()
If Not IsNull(toDate.value) Then
toDateH.value = ToHijriDate(toDate.value)
End If
End Sub

Private Sub ToDateH_LostFocus()
    VBA.Calendar = vbCalGreg
            toDate.value = ToGregorianDate(toDateH.value)
End Sub

Private Sub ToFiter_Change()
If Not IsNull(ToFiter.value) Then
ToFiterH.value = ToHijriDate(ToFiter.value)
End If
End Sub

Private Sub ToFiterH_LostFocus()
  VBA.Calendar = vbCalGreg
            ToFiter.value = ToGregorianDate(ToFiterH.value)
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

Private Sub VSFlexGrid1_Click()
If m_RetrunType = 8 Then
FrmOtheExpensAqar.FindRec val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("ID")))
ElseIf m_RetrunType = 808 Then
FrmCashing1.TxtContractNo.Text = (VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("ContNo")))
FrmCashing1.TxtContNo.Text = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("ID")))
End If
End Sub

Private Sub XPDtbBill_Change()
If Not IsNull(XPDtbBill.value) Then
         NourHijriCal1.value = ToHijriDate(XPDtbBill.value)
End If
End Sub
