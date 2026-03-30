VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMkafea 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„ﬂ«›√  "
   ClientHeight    =   4875
   ClientLeft      =   2685
   ClientTop       =   2475
   ClientWidth     =   6090
   HelpContextID   =   530
   Icon            =   "FrmMokfa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   6090
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1140
      Left            =   -15
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3135
      Width           =   6105
      _cx             =   10769
      _cy             =   2011
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
         Height          =   420
         Left            =   5235
         TabIndex        =   8
         Top             =   555
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   741
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
         ButtonImage     =   "FrmMokfa.frx":038A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   420
         Left            =   3255
         TabIndex        =   7
         Top             =   555
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   741
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
         ButtonImage     =   "FrmMokfa.frx":0724
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   420
         Left            =   4200
         TabIndex        =   9
         Top             =   555
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         ButtonImage     =   "FrmMokfa.frx":0ABE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   420
         Left            =   2205
         TabIndex        =   10
         Top             =   555
         Width           =   840
         _ExtentX        =   1482
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
         ButtonImage     =   "FrmMokfa.frx":0E58
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   420
         Left            =   1245
         TabIndex        =   11
         Top             =   555
         Width           =   735
         _ExtentX        =   1296
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
         ButtonImage     =   "FrmMokfa.frx":11F2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5010
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   1065
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
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
         ButtonImage     =   "FrmMokfa.frx":178C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   3960
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   1065
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
         ButtonImage     =   "FrmMokfa.frx":1B26
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   2940
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1095
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
         ButtonImage     =   "FrmMokfa.frx":1EC0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   420
         Left            =   165
         TabIndex        =   12
         Top             =   555
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
         ButtonImage     =   "FrmMokfa.frx":225A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   210
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   3930
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   870
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   2
         Left            =   4650
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   -15
      Width           =   6105
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   495
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3165
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Text            =   "modflag"
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtKhsmEdafa_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   90
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3690
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMokfa.frx":25F4
               Key             =   "Emp_Name"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMokfa.frx":298E
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMokfa.frx":2D28
               Key             =   "Emp_Code"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMokfa.frx":32C2
               Key             =   "Emp_Salary"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   16
         Top             =   75
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
         ButtonImage     =   "FrmMokfa.frx":365C
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   15
         Top             =   75
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
         ButtonImage     =   "FrmMokfa.frx":39F6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   14
         Top             =   75
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
         ButtonImage     =   "FrmMokfa.frx":3D90
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   13
         Top             =   75
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
         ButtonImage     =   "FrmMokfa.frx":412A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2160
         Picture         =   "FrmMokfa.frx":44C4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ﬂ«›√  "
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
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   90
         Width           =   2280
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2535
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   585
      Width           =   6090
      Begin VB.ComboBox CboCalType 
         Height          =   315
         Left            =   2475
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1260
         Width           =   2460
      End
      Begin VB.ComboBox CmbMonth 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3750
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2175
         Width           =   1305
      End
      Begin VB.ComboBox CmbYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2175
         Width           =   825
      End
      Begin VB.TextBox TxtNotes 
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
         Height          =   2100
         Left            =   60
         MaxLength       =   50
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   188
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   330
         Left            =   3360
         TabIndex        =   1
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   94044161
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DCEmp_Name 
         Height          =   315
         Left            =   2475
         TabIndex        =   2
         Top             =   930
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox TxtKhsmEdafa_Code 
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
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· —ﬁ„ «·„ﬂ«›√…"
         Top             =   188
         Width           =   1575
      End
      Begin VB.TextBox TxtKhsmEdafa_Value 
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
         Left            =   4125
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Text            =   "1"
         Top             =   1650
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„ﬂ«›√…"
         Height          =   255
         Index           =   5
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1290
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì‰›– ›Ï "
         Height          =   195
         Index           =   0
         Left            =   5385
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2235
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   195
         Index           =   1
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1980
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "”‰…"
         Height          =   195
         Index           =   2
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1980
         Width           =   270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "„·«ÕŸ« "
         Height          =   195
         Index           =   4
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   195
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌÊ„"
         Height          =   195
         Index           =   3
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1650
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·„ﬂ«›√…"
         Height          =   195
         Index           =   0
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„ﬂ«›√…"
         Height          =   195
         Index           =   6
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   255
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   195
         Index           =   0
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   975
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œ… «·„ﬂ«›√…"
         Height          =   255
         Index           =   1
         Left            =   4935
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   1110
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   42
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   4320
      Width           =   2460
      _ExtentX        =   4339
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
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ…"
      Height          =   270
      Index           =   13
      Left            =   2490
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   4350
      Width           =   915
   End
End
Attribute VB_Name = "FrmMkafea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim cSearch As clsDCboSearch
Dim RecId As String
Dim II As Long
Dim StrDate As Date

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
                           If ChekClodePeriod(DTPicker.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                      
                      
    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtKhsmEdafa_ID.Text <> "" Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbQuestion + vbMsgBoxRtlReading + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
            RsSavRec.find "KhsmEdafa_ID=" & val(TxtKhsmEdafa_ID.Text), , adSearchForward, 1
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
            '------------------------------ Move Next ---------------------------.
            BtnNext_Click
        End If
    
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
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
        FindRec val(TxtKhsmEdafa_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtKhsmEdafa_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
                           If ChekClodePeriod(DTPicker.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
    On Error GoTo ErrTrap
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtKhsmEdafa_ID.Text <> "" Then
        '        If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
        '            RsSavRec.MoveNext
        '            RsSavRec.MoveLast
        '        End If
        Me.DCUser.BoundText = user_id
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.DCEmp_Name.SetFocus
    End If

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"
    
    My_SQL = "select * From tblKhsmEdafa where KhsmEdafa_Type=1"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        TxtKhsmEdafa_Code.Text = rs.RecordCount + 1
    Else
        TxtKhsmEdafa_Code.Text = 1
    End If

    rs.Close
    CmbYear.Text = year(Date)
    CmbMonth.ListIndex = Month(Date) - 1
    TxtKhsmEdafa_Code.SetFocus
    Me.DCUser.BoundText = user_id
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtKhsmEdafa_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtKhsmEdafa_ID.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnSave_Click()
                           If ChekClodePeriod(DTPicker.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                      
                      
    On Error GoTo ErrTrap

    'Dim StrVacCode As String
    'Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Msg As String
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    If DCEmp_Name.BoundText = "" Then
       ' Msg = "⁄›Ê« Ì—ÃÏ  ÕœÌœ «”„ «·„ÊŸ›"
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
        Else
        Msg = "Select employee"
        End If


        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCEmp_Name.SetFocus
        Exit Sub
    End If

    StrDate = "1/" & CmbMonth.ListIndex + 1 & "/" & CmbYear.Text

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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtKhsmEdafa_ID.Text)
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

Private Sub CboCalType_Change()
    ChangeCalType
End Sub

Private Sub CboCalType_Click()
    ChangeCalType
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Premium/Bonus"
    Label1(2).Caption = Me.Caption

    Label1(4).Caption = "Remarks"

    Label1(6).Caption = "ID"
 
    Label2(0).Caption = "Date"
    Label1(0).Caption = "Employee"
    Label1(5).Caption = "premium Type"
    Label1(1).Caption = "Interval"
    Label1(3).Caption = "Day"
    Label3(0).Caption = "Start"
    Label3(1).Caption = "Month"
    Label3(2).Caption = "Year"

    Label2(2).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."

    Label1(13).Caption = "By"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

End Sub

Private Sub Form_Load()

    Dim BKGrndPic As ClsBackGroundPic
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap

    Set BKGrndPic = New ClsBackGroundPic

    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    My_SQL = "select * From tblKhsmEdafa where KhsmEdafa_Type=1 order by KhsmEdafa_Code"
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    Me.TxtModFlg.Text = "R"
    'load tblUsers -----------------------------------------------
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCUser
    Me.DCUser.BoundText = user_id
    'load tblEmployee -----------------------------------------------
    Dcombos.GetEmployees Me.DCEmp_Name
    Set cSearch = New clsDCboSearch
    Set cSearch.Client = DCEmp_Name

    For i = 1900 To 3000
        CmbYear.AddItem i
        CmbYear.ItemData(CmbYear.NewIndex) = i
    Next i


    CmbYear.Text = year(Date)
  
    SetDtpickerDate DTPicker
    Resize_Form Me


    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If





    If SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboCalType
            .Clear
            .AddItem "Days From Salary"
            .AddItem "Value"
        End With

    Else

        With Me.CboCalType
            .Clear
            .AddItem "„ﬂ«›√… √Ì«„ „‰ «·„— »"
            .AddItem "„ﬂ«›√… ﬁÌ„… ‰ﬁœÌ… ⁄·Ï «·„— »"
        End With

    End If
CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next
   CmbMonth.ListIndex = Month(Date) - 1
   
    
    BtnFirst_Click
    ShowTip

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
       
                'SaveData
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Terminate()
    Set cSearch = Nothing
    Set FrmVacancy = Nothing

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

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("tblKhsmEdafa", "KhsmEdafa_ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("KhsmEdafa_ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("KhsmEdafa_Type").value = 1
    RsSavRec.Fields("Emp_ID").value = IIf(DCEmp_Name.Text <> "", Trim(DCEmp_Name.BoundText), Null)
    RsSavRec.Fields("KhsmEdafa_Date").value = IIf(CStr(StrDate) <> "", StrDate, Null)
    RsSavRec.Fields("KhsmEdafa_Value").value = IIf(TxtKhsmEdafa_Value.Text <> "", Trim(TxtKhsmEdafa_Value.Text), Null)
    RsSavRec.Fields("KhsmEdafa_Code").value = IIf(TxtKhsmEdafa_Code.Text <> "", Trim(TxtKhsmEdafa_Code.Text), Null)
    RsSavRec.Fields("Notes").value = IIf(TxtNotes.Text <> "", Trim(TxtNotes.Text), Null)
    RsSavRec.Fields("RcDate").value = IIf(DTPicker.value <> "", Trim(DTPicker.value), Null)
    RsSavRec("CalculateValueType").value = IIf(Me.CboCalType.ListIndex = -1, 0, Me.CboCalType.ListIndex)
    RsSavRec("UserID").value = IIf(Me.DCUser.BoundText = "", user_id, Me.DCUser.BoundText)

    RsSavRec.update

    'MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If


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

    TxtKhsmEdafa_ID.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_ID").value), "", RsSavRec.Fields("KhsmEdafa_ID").value)
    DCEmp_Name.BoundText = IIf(IsNull(RsSavRec.Fields("Emp_ID").value), "", RsSavRec.Fields("Emp_ID").value)
    TxtKhsmEdafa_Code.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Code").value), "", RsSavRec.Fields("KhsmEdafa_Code").value)
    TxtKhsmEdafa_Value.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Value").value), "", RsSavRec.Fields("KhsmEdafa_Value").value)
    DTPicker.value = IIf(IsNull(RsSavRec.Fields("RcDate").value), Date, RsSavRec.Fields("RcDate").value)
    TxtNotes.Text = IIf(IsNull(RsSavRec.Fields("Notes").value), "", RsSavRec.Fields("Notes").value)
    CmbMonth.ListIndex = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Date").value), -1, Month(RsSavRec.Fields("KhsmEdafa_Date").value) - 1)
    CmbYear.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Date").value), "", year(RsSavRec.Fields("KhsmEdafa_Date").value))
    Me.CboCalType.ListIndex = IIf(IsNull(RsSavRec("CalculateValueType").value), 0, RsSavRec("CalculateValueType").value)
    Me.DCUser.BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub TxtKhsmEdafa_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap

    RsSavRec.find "KhsmEdafa_ID=" & RecId, , adSearchForward, 1

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

Private Sub DCEmp_Name_KeyPress(KeyAscii As Integer)
    KeyAscii = DataFormat(ChrOnly, KeyAscii)

End Sub

Private Sub TxtKhsmEdafa_Value_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtKhsmEdafa_Value.Text, 0)
End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
    
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
  
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtKhsmEdafa_ID.Text <> "" Then
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
    
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

Private Sub ChangeCalType()

    If CboCalType.ListIndex = -1 Then Exit Sub
    If SystemOptions.UserInterface = EnglishInterface Then
        If Me.CboCalType.ListIndex = 0 Then
            Me.Label1(1).Caption = "Interval"
            Me.Label1(3).Caption = "day"
        ElseIf Me.CboCalType.ListIndex = 1 Then
            Me.Label1(1).Caption = "value"
            Me.Label1(3).Caption = ""
        End If

    Else
        
        If Me.CboCalType.ListIndex = 0 Then
            Me.Label1(1).Caption = "„œ… «·„ﬂ«›√…"
            Me.Label1(3).Caption = "ÌÊ„"
        ElseIf Me.CboCalType.ListIndex = 1 Then
            Me.Label1(1).Caption = "ﬁÌ„… «·„ﬂ«›√…"
            Me.Label1(3).Caption = ""
        End If

    End If

End Sub

