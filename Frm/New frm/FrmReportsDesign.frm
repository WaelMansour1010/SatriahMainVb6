VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmReportsDesign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘…  ’„Ì„ «· ﬁ«Ì—"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "FrmReportsDesign.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   " ⁄œÌ· «· ’„Ì„"
      Height          =   495
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   6480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2085
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   11955
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "’·«ÕÌ«  «·„” Œœ„Ì‰"
         Height          =   1575
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   11895
         Begin VB.ComboBox DcbType1 
            Height          =   315
            ItemData        =   "FrmReportsDesign.frx":038A
            Left            =   240
            List            =   "FrmReportsDesign.frx":038C
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1320
            Visible         =   0   'False
            Width           =   4410
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            ItemData        =   "FrmReportsDesign.frx":038E
            Left            =   6240
            List            =   "FrmReportsDesign.frx":0390
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   720
            Width           =   4410
         End
         Begin VB.CheckBox ChAllBranch 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ· «·›—Ê⁄"
            Height          =   210
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtPath 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1680
            Visible         =   0   'False
            Width           =   4410
         End
         Begin MSDataListLib.DataCombo DcbUser 
            Height          =   330
            Left            =   6240
            TabIndex        =   27
            Top             =   360
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbGroup 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Visible         =   0   'False
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   1320
            TabIndex        =   29
            Top             =   360
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„” Œœ„ „⁄Ì‰"
            Height          =   255
            Index           =   0
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã„Ê⁄Â „⁄Ì‰Â"
            Height          =   255
            Index           =   4
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„œÌÊ·"
            Height          =   255
            Index           =   6
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›—⁄ „⁄Ì‰"
            Height          =   255
            Index           =   9
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«—"
            Height          =   255
            Index           =   10
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1680
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.TextBox TxtUnitID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9195
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   105
         Width           =   1425
      End
      Begin VB.TextBox TxtVacNamee 
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
         Height          =   345
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   105
         Width           =   3180
      End
      Begin VB.TextBox TxtVacName 
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
         Height          =   345
         Left            =   4920
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   105
         Width           =   3180
      End
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "FrmReportsDesign.frx":0392
         Left            =   2280
         List            =   "FrmReportsDesign.frx":03A2
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3150
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ «‰Ã·Ì“Ì"
         Height          =   375
         Index           =   1
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„"
         Height          =   285
         Index           =   3
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   90
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   255
         Index           =   0
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   90
         Width           =   1890
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   12015
      _cx             =   21193
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "‘«‘…  ’„Ì„ «· ﬁ«—Ì— "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
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
         ButtonImage     =   "FrmReportsDesign.frx":03BB
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   345
         Left            =   615
         TabIndex        =   3
         Top             =   150
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
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
         ButtonImage     =   "FrmReportsDesign.frx":0755
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   345
         Left            =   1065
         TabIndex        =   4
         Top             =   150
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
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
         ButtonImage     =   "FrmReportsDesign.frx":0AEF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   345
         Left            =   1530
         TabIndex        =   5
         Top             =   150
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
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
         ButtonImage     =   "FrmReportsDesign.frx":0E89
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton CmdInfo 
         Height          =   615
         Left            =   7560
         TabIndex        =   45
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmReportsDesign.frx":1223
         ButtonImageHover=   "FrmReportsDesign.frx":1EFD
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   990
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8490
      Width           =   11730
      _cx             =   20690
      _cy             =   1746
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12720
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   4800
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   570
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
         ButtonImage     =   "FrmReportsDesign.frx":2BD7
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   12
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
         ButtonImage     =   "FrmReportsDesign.frx":2F71
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   420
         Left            =   2355
         TabIndex        =   13
         Top             =   495
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
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
         ButtonImage     =   "FrmReportsDesign.frx":330B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ButPrient 
         Height          =   495
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â"
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
         ButtonImage     =   "FrmReportsDesign.frx":36A5
         ColorButton     =   14871017
         DisplayPersistentHover=   0   'False
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   135
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   135
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   135
         Width           =   975
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3645
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   12015
      _cx             =   21193
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmReportsDesign.frx":3A3F
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
   Begin ImpulseButton.ISButton btnNew 
      Height          =   420
      Left            =   8640
      TabIndex        =   37
      Top             =   6480
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   741
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«‰‘«¡  ﬁ—Ì—"
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
      ButtonImage     =   "FrmReportsDesign.frx":3B51
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton btnModify 
      Height          =   420
      Left            =   8640
      TabIndex        =   38
      Top             =   6825
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   741
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ· «· ﬁ—Ì—"
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
      ButtonImage     =   "FrmReportsDesign.frx":3EEB
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton btnDelete 
      Height          =   420
      Left            =   8640
      TabIndex        =   40
      Top             =   7920
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   741
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
      ButtonImage     =   "FrmReportsDesign.frx":4285
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton btnSave 
      Height          =   420
      Left            =   8640
      TabIndex        =   41
      Top             =   7200
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   741
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
      ButtonImage     =   "FrmReportsDesign.frx":481F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton BtnPrint 
      Height          =   525
      Left            =   12360
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   7560
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   926
      ButtonStyle     =   1
      ButtonPositionImage=   2
      Caption         =   "ÿ»«⁄Â «· ﬁ—Ì—"
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
      ButtonImage     =   "FrmReportsDesign.frx":4BB9
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton BtnUndo 
      Height          =   420
      Left            =   8640
      TabIndex        =   43
      Top             =   7560
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   741
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
      ButtonImage     =   "FrmReportsDesign.frx":4F53
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1935
      Left            =   120
      Top             =   6480
      Width           =   8415
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " „ﬂ‰ﬂ Â–… «·‘«‘… „‰ «‰‘«¡ «· ﬁ«—Ì— «·ÃœÌœ… »”ÂÊ·… ÊÌ”—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1860
      Index           =   5
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   6480
      Width           =   8295
   End
End
Attribute VB_Name = "FrmReportsDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim RecID As String
Dim II As Long

Private Sub ChangeLang()
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name Ar"
    Label1(1).Caption = "Name Eng"
    Frame1.Caption = "Users Validity"
    lbl(0).Caption = "User"
    lbl(6).Caption = "Type"
    lbl(10).Caption = "Path"
    lbl(4).Caption = "Group"
    lbl(9).Caption = "Branch"
    ChAllBranch.Caption = "All Branch"
    Command1.Caption = "Update"
    ChAllBranch.RightToLeft = False
'     Label1(2).Caption = " Remarks"
' Me.ButPrient.Caption = "Prient"
' ISButton1.Caption = "Search"
 btnQuery.Caption = "Search"
 
    With Grid
    .TextMatrix(0, .ColIndex("Ser")) = " Serial"
        .TextMatrix(0, .ColIndex("code")) = " Code"
        .TextMatrix(0, .ColIndex("name")) = " Name AR"
        .TextMatrix(0, .ColIndex("nameE")) = " Name Eng"
        .TextMatrix(0, .ColIndex("show")) = " Show"
        .TextMatrix(0, .ColIndex("path")) = "Path"
        .TextMatrix(0, .ColIndex("type")) = "Type"
        
        Me.Caption = "Design Reports"
        EleHeader.Caption = Me.Caption
        btnNew.Caption = "New"
        btnModify.Caption = "Modify"
        btnSave.Caption = "Save"
        BtnUndo.Caption = "Undo"
        btnDelete.Caption = "Delete"
        btnCancel.Caption = "Exit"
        Label2(0).Caption = "Current Record"
        Label2(1).Caption = "NO Of Record"
    End With

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
End Sub




Private Sub Command1_Click()
ShellExecute 0&, vbNullString, TxtPath.text, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
        Dim Dcombos As ClsDataCombos
 Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.DcbBranch
    Dcombos.GetUsers Me.DcbUser
    If SystemOptions.UserInterface = ArabicInterface Then
       DcbType.AddItem "Õ”«»« "
       DcbType.AddItem "«’Ê·"
       DcbType.AddItem "„»Ì⁄« "
       DcbType.AddItem "„‘ —Ì« "
       DcbType.AddItem "„Œ«“‰"
       DcbType.AddItem "„⁄«„·«  „«·ÌÂ"
       DcbType.AddItem "‘ƒÊ‰ «·„ÊŸ›Ì‰"
       DcbType.AddItem "≈‰ «Ã"
       DcbType.AddItem "’Ì«‰Â ⁄«„Â"
       DcbType.AddItem "’Ì«‰… ”Ì«—« "
       DcbType.AddItem "‰ﬁ·Ì« "
       DcbType.AddItem "„‘«—Ì⁄"
       
       DcbType1.AddItem "Õ”«»« "
       DcbType1.AddItem "«’Ê·"
       DcbType1.AddItem "„»Ì⁄« "
       DcbType1.AddItem "„‘ —Ì« "
       DcbType1.AddItem "„Œ«“‰"
       DcbType1.AddItem "„⁄«„·«  „«·ÌÂ"
       DcbType1.AddItem "‘ƒÊ‰ «·„ÊŸ›Ì‰"
       DcbType1.AddItem "≈‰ «Ã"
       DcbType1.AddItem "’Ì«‰Â ⁄«„Â"
       DcbType1.AddItem "’Ì«‰… ”Ì«—« "
       DcbType1.AddItem "‰ﬁ·Ì« "
       DcbType1.AddItem "„‘«—Ì⁄"
      
     Else
       DcbType.AddItem "Accounting"
       DcbType.AddItem "FixedAssets"
       DcbType.AddItem "Sales"
       DcbType.AddItem "Purchases"
       DcbType.AddItem "Stocks Control"
       DcbType.AddItem "Financial transactions"
       DcbType.AddItem "HR Management"
       DcbType.AddItem "Production"
       DcbType.AddItem "Maintenance"
       DcbType.AddItem "Maintenance Cars"
       DcbType.AddItem "Shipping and Distribution"
       DcbType.AddItem "Projects"
     
       DcbType1.AddItem "Accounting"
       DcbType1.AddItem "FixedAssets"
       DcbType1.AddItem "Sales"
       DcbType1.AddItem "Purchases"
       DcbType1.AddItem "Stocks Control"
       DcbType1.AddItem "Financial transactions"
       DcbType1.AddItem "HR Management"
       DcbType1.AddItem "Production"
       DcbType1.AddItem "Maintenance"
       DcbType1.AddItem "Maintenance Cars"
       DcbType1.AddItem "Shipping and Distribution"
       DcbType1.AddItem "Projects"
     End If
    
    ScreenNameArabic = "‘«‘…  ’„Ì„ «· ﬁ«—Ì— "
    ScreenNameEnglish = "Design Reports  "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

    Dim cGrdBack As New ClsBackGroundPic
    Set Me.Grid.WallPaper = cGrdBack.Picture
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String

    My_SQL = "TblDesignReport"
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect

    Me.TxtModFlg.text = "R"

    Resize_Form Me
    FillGridWithData

    With Me.Grid

        '.Cell(flexcpPicture, 0, .ColIndex("Dis_Count")) = Me.GrdImageList.ListImages("Dis_Count").ExtractIcon
        '.Cell(flexcpPicture, 0, .ColIndex("UnitName")) = Me.GrdImageList.ListImages("UnitName").ExtractIcon
        '.Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next

        .ExtendLastCol = True
    End With

    BtnFirst_Click
    ShowTip

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String

    If TxtUnitID.text <> "" Then
    
        'If UnitsHaveTransactions(val(TxtUnitID.text)) = True Then
        '    If SystemOptions.UserInterface = ArabicInterface Then
        '        Msg = " ·« Ì„ﬂ‰ Õ–› Â–… «·ÊÕœ… ·ÊÃÊœ ⁄„·Ì«  „— »ÿÂ »Â« "
        '    Else
        '        Msg = " Can't Modify Unit - Unit Have Transaction "
        '    End If
'
'            MsgBox Msg, vbCritical
'            Exit Sub
'        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Delete This Record", vbYesNo + vbMsgBoxRight, App.title)
        End If

        If MSGType = vbYes Then
            RsSavRec.find "id=" & val(TxtUnitID.text), , adSearchForward, 1

            If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
               'sa CuurentLogdata ("D")
               Kill TxtPath.text
 
                RsSavRec.delete

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
                Else
                    MsgBox "Delete Success...", vbOKOnly + vbMsgBoxRight, App.title
                End If

                '------------------------------ Move Next ---------------------------.
                FillGridWithData
                BtnNext_Click
            End If
        End If
    
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259

            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "Sorry .. can't Delete this record , Reason : Data integrity"
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

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
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
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
    
                Msg = "Sorry.. this record Already Deleted" & Chr(13)
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
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
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & Chr(13)
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If TxtUnitID.text <> "" Then
        '        If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
        '            RsSavRec.MoveNext
        '            RsSavRec.MovePrevious
        '        End If
        
       ' If UnitsHaveTransactions(val(TxtUnitID.text)) = True Then
       '     If SystemOptions.UserInterface = ArabicInterface Then
       '         Msg = " ·« Ì„ﬂ‰  ⁄œÌ· Â–… «·ÊÕœ… ·ÊÃÊœ ⁄„·Ì«  „— »ÿÂ »Â« "
       '     Else
       '         Msg = " Can't Modify Unit - Unit Have Transaction "
       '     End If

       '     MsgBox Msg, vbCritical
       '     Exit Sub
       ' End If

        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
       'sa CuurentLogdata
    End If

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147467259

            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & Chr(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & Chr(13)
                Msg = Msg & " Can't Edit this record now - Another user work with it now" & Chr(13)
       
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    On Error GoTo ErrTrap
    Dim My_SQL As String

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.text = "N"

    My_SQL = "TblDesignReport"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtUnitID.text = rs.RecordCount + 1
    Else
        TxtUnitID.text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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

    '---------------------- check if data Vaclete -----------------------
    If Trim(Me.TxtVacName.text) = "" Then
        Msg = "ÌÃ» ﬂ «»… «”„ «·‰Ê⁄ ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtVacName.SetFocus
        Exit Sub
    End If
 '    If Trim(Me.TxtCode.text) = "" Then
 '       Msg = "ÌÃ» ﬂ «»… ﬂÊœ «·‰Ê⁄ ...!!!"
 '       MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '       TxtCode.SetFocus
 '       Exit Sub
 '   End If
StrVacName = ""
    '------------------------------ check if Empcode exist ----------------------
    StrVacName = IsRecExist("TblDesignReport", "name", Trim(TxtVacName.text), "name", "ID<>'" & Trim(TxtUnitID.text) & "'")

    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "this Unit Already Exist"
        End If

        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
        Exit Sub
    End If
 '   StrVacName = IsRecExist("TblModelss", "code", Trim(Me.TxtCode.text), "code", "ID<>'" & Trim(TxtUnitID.text) & "'")

 '   If StrVacName <> "" Then
 '       If SystemOptions.UserInterface = ArabicInterface Then
 '           Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·ﬂÊœ „‰ ﬁ»·"
 '       Else
 '           Msg = "this Unit Already Exist"
 '       End If
'
'        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
'        TxtVacName.SetFocus
'        Exit Sub
'    End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
            
            
            
    End Select
ShellExecute 0&, vbNullString, TxtPath.text, vbNullString, vbNullString, vbNormalFocus
    Exit Sub
ErrTrap:

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Error in Enterd data", vbOKOnly + vbMsgBoxRight, App.title
    End If

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtUnitID.text)
    Me.TxtModFlg.text = "R"
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
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                btnSave_Click

                ' SaveData
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
   ' Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblDesignReport", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap
   Dim path As String
    RsSavRec.Fields("name").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
    If Me.TxtModFlg.text = "N" Then
    
    Copyfiles path, TxtUnitID.text, DcbType.ListIndex
    
    TxtPath.text = path
    End If
    RsSavRec.Fields("Path").value = IIf(Me.TxtPath.text <> "", Trim(TxtPath.text), Null)
    RsSavRec.Fields("UserID").value = IIf(val(Me.DcbUser.BoundText) <> 0, val(Me.DcbUser.BoundText), Null)
    RsSavRec.Fields("BranchID").value = IIf(val(Me.DcbBranch.BoundText) <> 0, val(Me.DcbBranch.BoundText), Null)
    RsSavRec.Fields("GroupID").value = IIf(val(Me.DcbGroup.BoundText) <> 0, val(Me.DcbGroup.BoundText), Null)
    RsSavRec.Fields("TypeID").value = IIf(val(Me.DcbType.ListIndex) <> -1, val(Me.DcbType.ListIndex), Null)
   If ChAllBranch.value = vbChecked Then
   RsSavRec.Fields("TypeID").value = 1
   Else
   RsSavRec.Fields("TypeID").value = 0
   End If
    RsSavRec.update
    
    If SystemOptions.UserInterface = ArabicInterface Then
     '   MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    '    MsgBox "Saved Successfully", vbOKOnly + vbMsgBoxRight, App.title
    End If

    FillGridWithData
    
    TxtModFlg = "R"
    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtUnitID.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    TxtPath.text = IIf(IsNull(RsSavRec.Fields("Path").value), "", RsSavRec.Fields("Path").value)
    Me.DcbUser.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), 0, RsSavRec.Fields("BranchID").value)
    Me.DcbGroup.BoundText = IIf(IsNull(RsSavRec.Fields("GroupID").value), 0, RsSavRec.Fields("GroupID").value)
    Me.DcbType.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeID").value), -1, RsSavRec.Fields("TypeID").value)
    If (RsSavRec.Fields("AllBranch").value) = True Then
    Me.ChAllBranch.value = vbChecked
    Else
    Me.ChAllBranch.value = vbUnchecked
    End If
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtUnitID.text) = .TextMatrix(i, .ColIndex("UnitID")) Then
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecID As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Grid
 Select Case .ColKey(Col)
                 Case "show"
                Dim FilePath As String
                 ' LngRow = Row
                 FilePath = .TextMatrix(Row, .ColIndex("path"))
                 print_report , FilePath

End Select
End With
End Sub
Function print_report(Optional NoteSerial As String, Optional StrFileName As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    'Dim StrFileName As String
    Dim Msg As String

' MySQL = " SELECT     dbo.TblItems.ItemID, dbo.TblItemDiamonds.type, dbo.TblItemDiamonds.unite, dbo.TblItemDiamonds.weight, dbo.TblItemDiamonds.indexe, dbo.TblItemDiamonds.Gestonf, dbo.TblItemDiamonds.color, dbo.TblItemDiamonds.quality"
'MySQL = MySQL & " FROM         dbo.TblItems INNER JOIN"
' MySQL = MySQL & "      dbo.TblItemDiamonds ON dbo.TblItems.ItemID = dbo.TblItemDiamonds.ItemID"
'MySQL = MySQL & " Where (dbo.TblItems.ItemID = " & val(XPTxtID.text) & ")"


'MySQL = MySQL & " Where (dbo.TblCommisRece.id =" & val(XPTxtID.text) & ")"

 

 
   
' If SystemOptions.UserInterface = ArabicInterface Then
'          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepItemDiamondG.rpt"
'     Else
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepItemDiamondG.rpt"
'       End If
'
'
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    If DcbType.ListIndex = 0 Then
    MySQL = "select * from ACCOUNTS"
   ElseIf DcbType.ListIndex = 1 Then
    MySQL = "select * from FixedAssets"
    End If
    
   RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    xReport.reporttitle = TxtVacName.text
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

   'sa RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("UnitID")))
ErrTrap:
End Sub

Sub Copyfiles(Optional ByRef path As String, Optional namerep As String, Optional NO As Integer)
Dim StrFileName As String
Dim StrFileName1 As String
 namerep = namerep & ".rpt"
 If NO = 0 Then
StrFileName = App.path & "\REPORTS\NewDisgn\" & "Accounts.rpt"
ElseIf NO = 1 Then
StrFileName = App.path & "\REPORTS\NewDisgn\" & "FixedAsset.rpt"
Else
StrFileName = App.path & "\REPORTS\NewDisgn\" & "new.rpt"

End If

StrFileName1 = App.path & "\REPORTS\AfterDisgn\" & namerep

FileCopy StrFileName, StrFileName1
path = StrFileName1
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
 Select Case .ColKey(Col)
Case "show"
 .ColComboList(.ColIndex("show")) = "..."
 End Select
End With
End Sub

'Private Sub TxtDis_Count_KeyPress(KeyAscii As Integer)
'    KeyAscii = DataFormat(CurOnly, KeyAscii)
'End Sub




Private Sub TxtUnitID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecID, , adSearchForward, 1

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

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
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
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtUnitID.text <> "" Then
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
    ElseIf TxtModFlg.text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
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
    My_SQL = "select * From TblDesignReport order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs.Fields("nameE").value), "", rs.Fields("nameE").value)
                .TextMatrix(i, .ColIndex("path")) = IIf(IsNull(rs.Fields("Path").value), "", rs.Fields("Path").value)
                DcbType1.ListIndex = IIf(IsNull(rs.Fields("TypeID").value), -1, rs.Fields("TypeID").value)
                .TextMatrix(i, .ColIndex("type")) = DcbType1.text
                
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
    Wrap = Chr(13) + Chr(10)

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
        If Me.TxtModFlg.text = "R" Then
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

'Function CuurentLogdata(Optional Currentmode As String)
'
'    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·ÊÕœ…   " & TxtUnitID.text & Chr(13) & "  «”„ «·ÊÕœ… " & TxtVacName.text
'
'    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Unit No   " & TxtUnitID.text & Chr(13) & " Unit Name" & TxtVacNamee.text
'
'    If Currentmode <> "D" Then
'        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
'    Else
'        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
'    End If
    
'End Function

Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
