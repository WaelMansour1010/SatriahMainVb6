VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmAsest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ⁄—Ì› «·„ÊÃÊœ«  «·⁄Ì‰Ì…"
   ClientHeight    =   7680
   ClientLeft      =   3570
   ClientTop       =   2025
   ClientWidth     =   10110
   Icon            =   "FrmAsest.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   9810
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   6
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   45
            Width           =   855
         End
      End
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
               Picture         =   "FrmAsest.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsest.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":245A
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   11
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":27F4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   12
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":2B8E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   13
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":2F28
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄—Ì› «·„ÊÃÊœ«  «·⁄Ì‰Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   90
         Width           =   3240
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2595
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4080
      Width           =   10995
      Begin VB.TextBox txtPrice 
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
         Left            =   6870
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Tag             =   "Õœœ «”„ «·⁄Âœ…"
         Top             =   1530
         Width           =   1635
      End
      Begin VB.TextBox TxtAsestName 
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
         Left            =   480
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Tag             =   "Õœœ «”„ «·⁄Âœ…"
         Top             =   840
         Width           =   8025
      End
      Begin VB.TextBox Txtname 
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
         Left            =   480
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Tag             =   "Õœœ «”„ «·⁄Âœ…"
         Top             =   480
         Width           =   8025
      End
      Begin VB.TextBox txtremark 
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
         Height          =   675
         Left            =   480
         MaxLength       =   300
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1860
         Width           =   8025
      End
      Begin VB.TextBox txtcode 
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
         Left            =   5970
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo Dcasest 
         Bindings        =   "FrmAsest.frx":32C2
         Height          =   315
         Left            =   480
         TabIndex        =   29
         Top             =   840
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   556
         _Version        =   393216
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
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   35
         Top             =   120
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "—»ÿ „‰ „·› «·«’Ê·"
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   36
         Top             =   120
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ÌœÊÌ"
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBrnch 
         Height          =   315
         Left            =   480
         TabIndex        =   38
         Top             =   1200
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”⁄—"
         Height          =   195
         Index           =   6
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1590
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   195
         Index           =   5
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
         Height          =   195
         Index           =   1
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«’·"
         Height          =   195
         Index           =   4
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄Âœ"
         Height          =   285
         Index           =   0
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ  "
         Height          =   195
         Index           =   3
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   870
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   -60
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6705
      Width           =   9840
      _cx             =   17357
      _cy             =   1799
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   1
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
      Begin ImpulseButton.ISButton btnNew 
         Height          =   330
         Left            =   6135
         TabIndex        =   16
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÃœÌœ"
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
         ButtonImage     =   "FrmAsest.frx":33CB
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   4590
         TabIndex        =   17
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ›Ÿ"
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
         ButtonImage     =   "FrmAsest.frx":3765
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   5355
         TabIndex        =   18
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ⁄œÌ·"
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
         ButtonImage     =   "FrmAsest.frx":3AFF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   3825
         TabIndex        =   19
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " —«Ã⁄"
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
         ButtonImage     =   "FrmAsest.frx":3E99
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   3060
         TabIndex        =   20
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–›"
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
         ButtonImage     =   "FrmAsest.frx":4233
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   90
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14737632
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":47CD
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   105
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14871017
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":4B67
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   4725
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         ButtonStyle     =   1
         ButtonPositionImage=   2
         Caption         =   ""
         BackColor       =   14871017
         FontSize        =   14.25
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAsest.frx":4F01
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   2265
         TabIndex        =   24
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "FrmAsest.frx":529B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   2505
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   225
         Width           =   975
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   540
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   3405
      Left            =   0
      TabIndex        =   34
      Top             =   600
      Width           =   9765
      _cx             =   17224
      _cy             =   6006
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmAsest.frx":5635
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
End
Attribute VB_Name = "FrmAsest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
     Dim StrSQL As String
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    'If TxtVac_ID.text <> "" Then
    '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
    '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    Else
        MSGType = MsgBox("Are you sure want to delete this recored", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    End If

    If MSGType = vbYes Then
     StrSQL = "Delete From FixedAssets Where CarsDataID =" & val(Me.txtcode.Text) & "  And FlgCarNotFixed = 3"
              Cn.Execute StrSQL, , adExecuteNoRecords
        RsSavRec.Find "AsCode=" & val(txtcode.Text), , adSearchForward, 1
        RsSavRec.delete
        
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Recored deleted successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        '------------------------------ Move Next ---------------------------.
        FillGridWithData
        BtnNext_Click
    End If

    'End If
    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "Sorry , This recored can't be deleted for the association with other records "
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
            Msg = Msg & "Data will be updated now" & CHR(13)
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
            Msg = Msg & "Date will be updated now" & CHR(13)
        End If
       
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
     '   Me.TxtVacName.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & "This recored can't be edited now" & CHR(13)
                Msg = Msg & "it's under modification by other user on the network"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
FillGridWithData

    TxtModFlg.Text = "N"
Rd(0).value = True
Rd_Click (0)

    My_SQL = "TblAssestes"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
    '    TxtSerial.text = rs.RecordCount + 1
    Else
       ' TxtSerial.text = 1
    End If

    rs.Close
 '   CmbType.ListIndex = 0
   ' TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext

        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
            Msg = Msg & "Date will be updated now" & CHR(13)
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
            Msg = Msg & "Date will be updated now" & CHR(13)
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
'    If txtcode.text <> "" Then
    '---------------------- check if data Vaclete -----------------------

'    For Each CtrlTxt In Me.Controls
'
'        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
'            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
'                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
'                CtrlTxt.SetFocus
'                Exit Sub
'            End If
'        End If
'
'    Next
If Rd(0).value = False And Rd(1).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ Õ«·… «·—»ÿ "
Else
MsgBox "Please Select Type Link"
End If
Exit Sub
End If
If Rd(0).value = True Then
If Dcasest.Text = "" Or val(Dcasest.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·«’·"
Else
MsgBox "Please Select Assest"
End If
Exit Sub
End If
End If
If Rd(1).value = True Then
If TxtAsestName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·«’·"
Else
MsgBox "Please Eneter Assest"
End If
Exit Sub
End If
End If
    '------------------------------ check if Empcode exist ----------------------

   ' StrVacName = IsRecExist("TblAssestes", "AsName", Trim(Dcasest.text), "AsName", "Ascode<>'" & Trim(txtCode.text) & "'")

   ' If StrVacName <> "" Then
   '     Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
   '
   '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
   '
   '
   '     Exit Sub
'
'    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
Else
    MsgBox "Something is wrong in entered data", vbOKOnly + vbMsgBoxRight, App.title
End If
'    Else
'     MsgBox "Õﬁ· «·ﬂÊœ «Ã»«—Ì"
'End If

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click

    If FristCount = LastCount Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "There's no new data"
        End If
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        Else
            Msg = "No. of records befor update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        End If
        If LastCount > FristCount Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
 Dim Dcombos As ClsDataCombos
    My_SQL = "TblAssestes"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
     Set Dcombos = New ClsDataCombos
   Dcombos.GetFixedAssets Me.Dcasest
   Dcombos.GetBranches Me.DcbBrnch
    FillGridWithData

    With Me.Grid
   .Cell(flexcpPicture, 0, .ColIndex("AsName")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
     .Cell(flexcpPicture, 0, .ColIndex("AsCode")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
     .Cell(flexcpPicture, 0, .ColIndex("Fixed")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon

      For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
  End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    'Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

ErrTrap:
End Sub

Private Sub ChangeLang()

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic


    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    
    '###################### Khaled ######################
    Label1(2).Caption = "Declaration  of Assets"
    Me.Caption = Label1(2).Caption
    Label1(3).Caption = "Code"
    Rd(0).Caption = "Link From Asset"
    Rd(1).Caption = "Manual"
    Label1(0).Caption = "Asset Name"
    Label1(4).Caption = "Asset Name Eng"
    Label1(5).Caption = "Branch"
    Label1(1).Caption = "Notes"
    
    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "Recored Count"
    
    With Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("AsCode")) = "Asset Code"
        .TextMatrix(0, .ColIndex("AsName")) = "Asset Name"
        .TextMatrix(0, .ColIndex("Fixed")) = "Asset Name Eng"

    End With
    '####################################################

End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblAssestes", "AsID", "")
    txtcode.Text = StrRecID
    RsSavRec.AddNew
   RsSavRec.Fields("AsCode").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("AsName").value = IIf(Txtname.Text <> "", Trim(Txtname.Text), Null)
    RsSavRec.Fields("AsDes").value = IIf(txtremark.Text <> "", Trim(txtremark.Text), Null)
    RsSavRec.Fields("AsestName").value = (IIf(TxtAsestName.Text <> "", Trim(TxtAsestName.Text), Null))
    RsSavRec.Fields("AsFixedID").value = IIf(Dcasest.BoundText <> "", val(Dcasest.BoundText), Null)
    RsSavRec.Fields("BranchID").value = IIf(DcbBrnch.BoundText <> "", val(DcbBrnch.BoundText), Null)
    
    RsSavRec.Fields("Price").value = IIf(txtPrice.Text <> "", Trim(txtPrice.Text), Null)
   If Rd(1).value = True Then
   RsSavRec.Fields("TypeEqup").value = 1
   Else
   RsSavRec.Fields("TypeEqup").value = 0
   End If
    RsSavRec.update
If Rd(1).value = True Then
 SaveFixed
 End If
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGridWithData
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub
Sub SaveFixed()
Dim sql As String
Dim StrSQL As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim FexdID As Double

If Me.TxtModFlg.Text = "N" Then
sql = "Select * from FixedAssets where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
FexdID = CStr(new_id("FixedAssets", "id", "", True))
Rs5.AddNew
Else
FexdID = val(Dcasest.BoundText)
sql = "Select * from FixedAssets where id=" & FexdID & ""
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
End If
Rs5("id").value = FexdID
Rs5("CarsDataID").value = val(txtcode.Text)
Rs5("FlgCarNotFixed").value = 3
Rs5("ISEQUP").value = 1
Rs5("HaveDepreciation").value = 1
Rs5("PurchasePrice").value = 1
Rs5("Name").value = TxtAsestName.Text
Rs5("code").value = txtcode.Text
Rs5("branch_no").value = val(DcbBrnch.BoundText)

Rs5.update
sql = "Update TblAssestes set AsFixedID=" & FexdID & "  where AsCode =" & val(txtcode.Text) & ""
Cn.Execute sql

End Sub
Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("AsID").value), "", RsSavRec.Fields("AsID").value)
  Txtname.Text = IIf(IsNull(RsSavRec.Fields("AsName").value), "", RsSavRec.Fields("AsName").value)
    Me.txtremark.Text = IIf(IsNull(RsSavRec.Fields("AsDes").value), "", RsSavRec.Fields("AsDes").value)
    Me.txtcode.Text = IIf(IsNull(RsSavRec.Fields("AsCode").value), "", RsSavRec.Fields("AsCode").value)
    Me.TxtAsestName.Text = IIf(IsNull(RsSavRec.Fields("AsestName").value), "", RsSavRec.Fields("AsestName").value)
    Me.Dcasest.BoundText = val(IIf(IsNull(RsSavRec.Fields("AsFixedID").value), "", RsSavRec.Fields("AsFixedID").value))
    Me.DcbBrnch.BoundText = val(IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value))
    Me.txtPrice.Text = IIf(IsNull(RsSavRec.Fields("Price").value), "", RsSavRec.Fields("Price").value)
  If IsNull(RsSavRec.Fields("TypeEqup").value) Then
    Rd(0).value = True
  Else
           If (RsSavRec.Fields("TypeEqup").value) = 1 Then
                Rd(1).value = True
            Else
                Rd(0).value = True
            End If
  End If
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

          If Trim(txtcode.Text) = .TextMatrix(i, .ColIndex("AsCode")) Then
               ' TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
               Txtname.Text = .TextMatrix(i, .ColIndex("AsName"))
               Me.Dcasest.Text = .TextMatrix(i, .ColIndex("Fixed"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("AsCode")))
ErrTrap:
End Sub


Private Sub Rd_Click(Index As Integer)
If Rd(1).value = True Then
Dcasest.Visible = False
TxtAsestName.Visible = True
ElseIf Rd(0).value = True Then
Dcasest.Visible = True
TxtAsestName.Visible = False
End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "AsCode=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
      Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
       Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = True

        If txtcode.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.Text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
      '  Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblAssestes order by AsID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
          rs.MoveFirst

           For i = 1 To .Rows - 1
          .TextMatrix(i, .ColIndex("Ser")) = i
             
              .TextMatrix(i, .ColIndex("AsName")) = IIf(IsNull(rs.Fields("AsName").value), "", rs.Fields("AsName").value)
          
                .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(rs.Fields("AsCode").value), "", rs.Fields("AsCode").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs.Fields("Price").value), "", rs.Fields("Price").value)
              Me.Dcasest.BoundText = IIf(IsNull(rs.Fields("AsFixedID").value), "", rs.Fields("AsFixedID").value)
                .TextMatrix(i, .ColIndex("Fixed")) = Me.Dcasest.Text
            
                rs.MoveNext
           Next

            rs.Close
      End If

        .RowHeight(-1) = 300
  End With

ErrTrap:
End Sub

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If

    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If

    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If

    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If

    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If

    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If

    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If

    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If

    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If

    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If

    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If

    'End If

    Exit Sub
ErrTrap:
End Sub

Private Function CheckDelCountry(LngCountryID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where CountryID=" & LngCountryID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry = False
    Else
        CheckDelCountry = True
    End If

    rs.Close
    Set rs = Nothing
End Function

