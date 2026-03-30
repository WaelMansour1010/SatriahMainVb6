VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form frmProductLine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  ŒÿÊÿ «·«‰ «Ã"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   Icon            =   "frmProductLine.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   89
      Top             =   720
      Width           =   12135
      Begin VB.CheckBox chkIsBasicLine 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Œÿ «”«”Ì"
         Height          =   255
         Left            =   420
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   210
         Width           =   1935
      End
      Begin VB.TextBox TxtCode 
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
         Left            =   5400
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Tag             =   "⁄›Ê« Ì—ÃÌ «œŒ«· ﬂÊœ «·Œÿ"
         Top             =   165
         Width           =   2265
      End
      Begin VB.TextBox TxtSerial 
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
         Height          =   315
         Left            =   8880
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   165
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·Œÿ"
         Height          =   195
         Index           =   11
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·Œÿ"
         Height          =   195
         Index           =   3
         Left            =   10425
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   270
         Width           =   1470
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   -60
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5775
      Width           =   12240
      _cx             =   21590
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
         Left            =   10455
         TabIndex        =   12
         Top             =   195
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
         ButtonImage     =   "frmProductLine.frx":058A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   7350
         TabIndex        =   13
         Top             =   195
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
         ButtonImage     =   "frmProductLine.frx":0924
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   8955
         TabIndex        =   14
         Top             =   195
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
         ButtonImage     =   "frmProductLine.frx":0CBE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   6225
         TabIndex        =   15
         Top             =   195
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
         ButtonImage     =   "frmProductLine.frx":1058
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   5100
         TabIndex        =   16
         Top             =   195
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
         ButtonImage     =   "frmProductLine.frx":13F2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   2040
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   690
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
         ButtonImage     =   "frmProductLine.frx":198C
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   2685
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   705
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
         ButtonImage     =   "frmProductLine.frx":1D26
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   1725
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   630
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
         ButtonImage     =   "frmProductLine.frx":20C0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   3705
         TabIndex        =   20
         Top             =   195
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
         ButtonImage     =   "frmProductLine.frx":245A
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   225
         Width           =   975
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   540
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   12225
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   2
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
            TabIndex        =   3
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
               Picture         =   "frmProductLine.frx":27F4
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":2B8E
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":2F28
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":32C2
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":365C
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":39F6
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":3D90
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProductLine.frx":432A
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   6
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
         ButtonImage     =   "frmProductLine.frx":46C4
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   7
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
         ButtonImage     =   "frmProductLine.frx":4A5E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
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
         ButtonImage     =   "frmProductLine.frx":4DF8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
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
         ButtonImage     =   "frmProductLine.frx":5192
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰«  ŒÿÊÿ «·«‰ «Ã"
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
         Left            =   8655
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   90
         Width           =   3360
      End
   End
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   4665
      Left            =   0
      TabIndex        =   25
      Top             =   1290
      Width           =   12075
      _cx             =   21299
      _cy             =   8229
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   -2147483630
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "»Ì«‰«  «·Œÿ «·«”«”Ì…|  »Ì«‰«  «·„⁄œ« /«·„«ﬂÌ‰«  |»Ì«‰«  «·⁄«„·Ì‰|»Ì«‰«  ﬂ· «·ŒÿÊÿ|«·„” Œœ„Ì‰"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   6
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4575
         Index           =   1
         Left            =   45
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   45
         Width           =   10095
         _cx             =   17806
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   105
            Index           =   1
            Left            =   13185
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   3600
            Width           =   1320
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4530
            Index           =   0
            Left            =   0
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   10095
            _cx             =   17806
            _cy             =   7990
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
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
            Begin VB.Frame Frm2 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4005
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   255
               Width           =   9915
               Begin VB.ComboBox cmbFormPrint 
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   107
                  Top             =   2670
                  Width           =   2025
               End
               Begin VB.TextBox TxtHourdipp 
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
                  Left            =   5160
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   2760
                  Width           =   2280
               End
               Begin VB.ComboBox CmbType 
                  BackColor       =   &H80000018&
                  Height          =   315
                  ItemData        =   "frmProductLine.frx":552C
                  Left            =   2280
                  List            =   "frmProductLine.frx":553C
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   3270
                  Visible         =   0   'False
                  Width           =   1005
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
                  Height          =   315
                  Left            =   3840
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„   «·Œÿ"
                  Top             =   45
                  Width           =   3600
               End
               Begin VB.TextBox TXTUsedPowerPriceH 
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
                  Left            =   5160
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   1800
                  Width           =   2280
               End
               Begin VB.TextBox TXTUsedElectricPriceH 
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
                  Left            =   5160
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   2280
                  Width           =   2280
               End
               Begin VB.TextBox TXTNotes 
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
                  Height          =   735
                  Left            =   120
                  MaxLength       =   50
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   33
                  Top             =   600
                  Width           =   7320
               End
               Begin VB.TextBox TxtWorkerPriceH 
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
                  TabIndex        =   32
                  Top             =   1800
                  Width           =   2040
               End
               Begin VB.TextBox TxtLinePriceH 
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
                  TabIndex        =   31
                  Top             =   2280
                  Width           =   2040
               End
               Begin MSDataListLib.DataCombo DcShift 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   38
                  Tag             =   "«Œ — «·‘Ì›"
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcManger 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   39
                  Tag             =   "«Œ — „œÌ— «·Œÿ"
                  Top             =   0
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboStoreName 
                  Height          =   315
                  Left            =   4500
                  TabIndex        =   108
                  Top             =   1350
                  Width           =   2970
                  _ExtentX        =   5239
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Œ“‰"
                  Height          =   375
                  Index           =   50
                  Left            =   9090
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1320
                  Width           =   600
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰„«–Ã «·ÿ»«⁄…"
                  Height          =   555
                  Index           =   18
                  Left            =   2310
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   2760
                  Width           =   2190
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·«Â·«ﬂ ›Ì «·”«⁄…"
                  Height          =   435
                  Index           =   16
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   2760
                  Width           =   2430
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œÿ"
                  Height          =   285
                  Index           =   0
                  Left            =   7860
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   120
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «” Â·«ﬂ «·ÊﬁÊœ ›Ì «·”«⁄Â"
                  Height          =   435
                  Index           =   1
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1800
                  Width           =   2190
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «” Â·«ﬂ «·ﬂÂ—»«¡›Ì «·”«⁄Â"
                  Height          =   555
                  Index           =   4
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   2280
                  Width           =   2430
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   285
                  Index           =   5
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   720
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·Â ›Ì «·”«⁄Â"
                  Height          =   555
                  Index           =   9
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   1800
                  Width           =   2190
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «” Â·«ﬂ «·Œÿ ›Ì «·”«⁄Â"
                  Height          =   555
                  Index           =   10
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   2400
                  Width           =   2190
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‘Ì› "
                  Height          =   285
                  Index           =   12
                  Left            =   7800
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ— «·Œÿ"
                  Height          =   285
                  Index           =   14
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   0
                  Width           =   1050
               End
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   330
               Index           =   0
               Left            =   12930
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   3480
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄·Ìﬁ:"
               Height          =   120
               Index           =   0
               Left            =   11610
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Tag             =   "22"
               Top             =   405
               Width           =   3135
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Code"
               Height          =   420
               Index           =   6
               Left            =   11370
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   3480
               Width           =   1440
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Height          =   360
               Index           =   2
               Left            =   6645
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   4965
               Width           =   8160
            End
         End
         Begin VB.Label Lb_note_value_by_characters 
            Alignment       =   1  'Right Justify
            Height          =   570
            Left            =   6765
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   4875
            Width           =   7860
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   285
            Left            =   11430
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   3600
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄·Ìﬁ:"
            Height          =   165
            Index           =   6
            Left            =   11610
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Tag             =   "22"
            Top             =   300
            Width           =   2895
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4575
         Index           =   2
         Left            =   12720
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   45
         Width           =   10095
         _cx             =   17806
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
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
         Begin VB.TextBox TxtHourdippTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   3510
            Width           =   2235
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   2
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   3570
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox TxtEquiomentPowerTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5310
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   3510
            Width           =   2235
         End
         Begin VB.TextBox TxtEquiomentElectricTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   3510
            Width           =   2235
         End
         Begin VSFlex8Ctl.VSFlexGrid Gridx 
            Height          =   3075
            Left            =   0
            TabIndex        =   58
            Top             =   360
            Width           =   9975
            _cx             =   17595
            _cy             =   5424
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
            Rows            =   1
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmProductLine.frx":5555
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
            ExplorerBar     =   0
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
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   375
            Left            =   60
            TabIndex        =   59
            Tag             =   "Delete Row"
            Top             =   3990
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmProductLine.frx":56BC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Height          =   480
            Index           =   3
            Left            =   1335
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   4965
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   390
            Index           =   7
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   3540
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄·Ìﬁ:"
            Height          =   150
            Index           =   1
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Tag             =   "22"
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            Height          =   255
            Left            =   8025
            TabIndex        =   60
            Top             =   3510
            Width           =   1575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4575
         Index           =   3
         Left            =   13020
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   45
         Width           =   10095
         _cx             =   17806
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
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
         Begin VB.Frame Frame10 
            Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·Œÿ"
            Height          =   4230
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   120
            Width           =   10035
            Begin VB.TextBox txt_emp_salary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   3
               Left            =   240
               TabIndex        =   74
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox txt_emp_salary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   2
               Left            =   1200
               TabIndex        =   73
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox txt_emp_salary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   1
               Left            =   2160
               TabIndex        =   72
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox txt_employee_count 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   3
               Left            =   240
               TabIndex        =   71
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox txt_employee_count 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   2
               Left            =   1200
               TabIndex        =   70
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox txt_employee_count 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   1
               Left            =   2160
               TabIndex        =   69
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox txt_employee_count 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   0
               Left            =   3000
               TabIndex        =   68
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox txt_emp_salary 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   360
               Index           =   0
               Left            =   3000
               TabIndex        =   67
               Top             =   3240
               Width           =   735
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   2340
               Left            =   120
               TabIndex        =   75
               Top             =   360
               Width           =   9720
               _cx             =   17145
               _cy             =   4128
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
               Rows            =   1
               Cols            =   13
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmProductLine.frx":56D8
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
               ExplorerBar     =   0
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
            Begin ALLButtonS.ALLButton opr_emplyees_name 
               Height          =   375
               Left            =   3840
               TabIndex        =   76
               Top             =   3600
               Visible         =   0   'False
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Õ›Ÿ «·⁄„«·…"
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
               BCOL            =   16711680
               BCOLO           =   16711680
               FCOL            =   16777215
               FCOLO           =   0
               MCOL            =   192
               MPTR            =   1
               MICON           =   "frmProductLine.frx":58A8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton CmdRemove 
               Height          =   375
               Left            =   240
               TabIndex        =   77
               Tag             =   "Delete Row"
               Top             =   3720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Õ–› ”ÿ—"
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
               BCOL            =   0
               BCOLO           =   0
               FCOL            =   255
               FCOLO           =   255
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmProductLine.frx":58C4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton ALLButton2 
               Height          =   375
               Left            =   7560
               TabIndex        =   94
               Top             =   2760
               Visible         =   0   'False
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   " ÕœÌÀ ”«⁄«  «·⁄„·"
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
               BCOL            =   16711680
               BCOLO           =   16711680
               FCOL            =   16777215
               FCOLO           =   0
               MCOL            =   192
               MPTR            =   1
               MICON           =   "frmProductLine.frx":58E0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ã„«·Ì ⁄œœ «·⁄„«·"
               Height          =   255
               Left            =   4200
               TabIndex        =   79
               Top             =   2880
               Width           =   1815
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·Â"
               Height          =   255
               Left            =   4200
               TabIndex        =   78
               Top             =   3240
               Width           =   1815
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   105
            Index           =   3
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   3630
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Height          =   570
            Index           =   4
            Left            =   1335
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   4905
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   285
            Index           =   8
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   3630
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄·Ìﬁ:"
            Height          =   165
            Index           =   2
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Tag             =   "22"
            Top             =   300
            Width           =   600
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4575
         Index           =   4
         Left            =   13320
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   45
         Width           =   10095
         _cx             =   17806
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   195
            Index           =   4
            Left            =   13110
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   3570
            Width           =   1395
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   2325
            Left            =   120
            TabIndex        =   85
            Top             =   630
            Width           =   9435
            _cx             =   16642
            _cy             =   4101
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmProductLine.frx":58FC
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
            ExplorerBar     =   0
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄·Ìﬁ:"
            Height          =   150
            Index           =   3
            Left            =   11190
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Tag             =   "22"
            Top             =   240
            Width           =   2955
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   270
            Index           =   15
            Left            =   10995
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   2790
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   435
            Index           =   5
            Left            =   6465
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   3855
            Width           =   7740
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4575
         Index           =   5
         Left            =   13620
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   45
         Width           =   10095
         _cx             =   17806
         _cy             =   8070
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   195
            Index           =   5
            Left            =   13110
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   3570
            Width           =   1395
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   3240
            Left            =   150
            TabIndex        =   104
            Top             =   150
            Width           =   9720
            _cx             =   17145
            _cy             =   5715
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
            Rows            =   1
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmProductLine.frx":5A0B
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
            ExplorerBar     =   0
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
         Begin ALLButtonS.ALLButton CmdRemove2 
            Height          =   375
            Left            =   240
            TabIndex        =   105
            Tag             =   "Delete Row"
            Top             =   3450
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmProductLine.frx":5B31
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   435
            Index           =   6
            Left            =   6465
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   3855
            Width           =   7740
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   270
            Index           =   17
            Left            =   10995
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   2790
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄·Ìﬁ:"
            Height          =   150
            Index           =   4
            Left            =   11190
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Tag             =   "22"
            Top             =   240
            Width           =   2955
         End
      End
   End
End
Attribute VB_Name = "frmProductLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim StrSQL  As String


Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If
    'ff

End Sub

Private Sub ALLButton1_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
     
    If Gridx.Rows > 1 Then
        If Gridx.Rows = 2 Then
            Me.Gridx.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Gridx.Rows > 1 Then
                If Me.Gridx.Row <> Me.Gridx.FixedRows - 1 Then
                    Me.Gridx.RemoveItem (Me.Gridx.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

End Sub

Private Sub ALLButton2_Click()
Dim i As Integer
Dim EmployeeSalary As Double
Dim Emp_id As Integer
    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1
'code
            If .TextMatrix(i, .ColIndex("Emp_id")) <> "" Or .TextMatrix(i, .ColIndex("code")) <> "" Then
               Emp_id = val(.TextMatrix(i, .ColIndex("Emp_id")))
               If Emp_id = 0 Then
               GetEmployeeIDFromCode .TextMatrix(i, .ColIndex("code")), Emp_id
               .TextMatrix(i, .ColIndex("Emp_id")) = Emp_id
               End If
               
                 EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_id"))), "")
 .TextMatrix(.Row, .ColIndex("hourprice")) = Round(EmployeeSalary / 240, 2)


        
            End If

        Next i
 
    End With

End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    'If TxtVac_ID.text <> "" Then
    '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
    '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '    End If
    MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

    If MSGType = vbYes Then
        RsSavRec.Find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
        CuurentLogdata ("D")
        RsSavRec.delete
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        '------------------------------ Move Next ---------------------------.
        FillGridWithData
        BtnNext_Click
    End If

    'End If
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
    
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        VSFlexGrid1.Enabled = True
      
        VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
        VSFlexGrid2.Enabled = True
      
        Gridx.Rows = Gridx.Rows + 1
        Gridx.Enabled = True
      
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

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
        
    IntCounter = 0
    Dim shift1_no_of_employee As Integer
    Dim shift2_no_of_employee As Integer
    Dim shift3_no_of_employee As Integer
    Dim shift4_no_of_employee As Integer

    Dim shift1_no_of_employee_Cost As Double
    Dim shift2_no_of_employee_Cost As Double
    Dim shift3_no_of_employee_Cost As Double
    Dim shift4_no_of_employee_Cost As Double

    shift1_no_of_employee_Cost = 0
    shift2_no_of_employee_Cost = 0
    shift3_no_of_employee_Cost = 0
    shift4_no_of_employee_Cost = 0

    shift1_no_of_employee = 0
    shift2_no_of_employee = 0
    shift3_no_of_employee = 0
    shift4_no_of_employee = 0

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift1")) = flexChecked Then
                    shift1_no_of_employee = shift1_no_of_employee + 1
                    shift1_no_of_employee_Cost = shift1_no_of_employee_Cost + val(.TextMatrix(i, .ColIndex("hourprice")))
                End If
        
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift2")) = flexChecked Then
                    shift2_no_of_employee = shift2_no_of_employee + 1
                    shift2_no_of_employee_Cost = shift2_no_of_employee_Cost + val(.TextMatrix(i, .ColIndex("hourprice")))
  
                End If
  
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift3")) = flexChecked Then
                    shift3_no_of_employee = shift3_no_of_employee + 1
                    shift3_no_of_employee_Cost = shift3_no_of_employee_Cost + val(.TextMatrix(i, .ColIndex("hourprice")))
  
                End If
  
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift4")) = flexChecked Then
                    shift4_no_of_employee = shift4_no_of_employee + 1
                    shift4_no_of_employee_Cost = shift4_no_of_employee_Cost + val(.TextMatrix(i, .ColIndex("hourprice")))
                End If
        
            End If

        Next i
 
    End With

    txt_employee_count(0).Text = shift1_no_of_employee
    txt_employee_count(1).Text = shift2_no_of_employee
    txt_employee_count(2).Text = shift3_no_of_employee
    txt_employee_count(3).Text = shift4_no_of_employee
    txt_emp_salary(0).Text = shift1_no_of_employee_Cost
    txt_emp_salary(1).Text = shift2_no_of_employee_Cost
    txt_emp_salary(2).Text = shift3_no_of_employee_Cost
    txt_emp_salary(3).Text = shift4_no_of_employee_Cost

    With Me.Gridx
        Me.TxtEquiomentPowerTotal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("UsedPowerPriceH"), .Rows - 1, .ColIndex("UsedPowerPriceH"))
        Me.TxtEquiomentElectricTotal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("UsedElectricPriceH"), .Rows - 1, .ColIndex("UsedElectricPriceH"))
        Me.TxtHourdippTotal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Hourdipp"), .Rows - 1, .ColIndex("Hourdipp"))
    End With

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
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
          
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    VSFlexGrid2.Enabled = True
             
    
    Gridx.Clear flexClearScrollable, flexClearEverything
    Gridx.Rows = 2
    Gridx.Enabled = True
          
    TxtModFlg.Text = "N"

    My_SQL = "TblProductLine"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    rs.Close
    CmbType.ListIndex = 0
    TXTCode.SetFocus
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

    TxtModFlg = "R"

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

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
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
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

    If Trim(DCboStoreName.Text) = "" Then

        If SystemOptions.UserInterface = ArabicInterface Then
        
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ «·Œÿ ﬁ»· «œŒ«· „Œ“‰ «·’—›"
        Else
            Msg = "The line can not be saved before inserting the exchange store"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If


    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblProductLine", "name", Trim(TxtVacName.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·Œÿ „‰ ﬁ»·"
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

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
    FindRec val(TxtSerial.Text)
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

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
     
    If VSFlexGrid1.Rows > 1 Then
        If VSFlexGrid1.Rows = 2 Then
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.VSFlexGrid1.Rows > 1 Then
                If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                    Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

End Sub
Private Sub CmdRemove2_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
     
    If VSFlexGrid2.Rows > 1 Then
        If VSFlexGrid2.Rows = 2 Then
            Me.VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.VSFlexGrid2.Rows > 1 Then
                If Me.VSFlexGrid2.Row <> Me.VSFlexGrid2.FixedRows - 1 Then
                    Me.VSFlexGrid2.RemoveItem (Me.VSFlexGrid2.Row)
                End If
            End If
        End If
    End If
            
   'ReLineGrid

End Sub
Function loadcombo()
    Dim My_SQL As String

    My_SQL = "select SeftCode,SheftName From TbLSheft "
    fill_combo DcShift, My_SQL

    My_SQL = "select Emp_ID,Emp_Name From TblEmployee "
    fill_combo Me.DcManger, My_SQL

End Function

Private Sub DcManger_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        loadcombo
    End If

End Sub

Private Sub DcShift_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF5 Then
        loadcombo
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    
    Dim Dcombos As New ClsDataCombos
    ScreenNameArabic = "»Ì«‰«  ŒÿÊÿ «·«‰ «Ã"
    ScreenNameEnglish = "Production Line Data"""
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
 
    
'cmbFormPrint.lis

cmbFormPrint.Clear
        
    For i = 1 To 10
        cmbFormPrint.AddItem CStr(i)
    Next
 
    My_SQL = "TblProductLine"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL



    My_SQL = "select SeftCode,SheftName From TbLSheft "
    fill_combo DcShift, My_SQL

    My_SQL = "select Emp_ID,Emp_Name From TblEmployee "
    fill_combo Me.DcManger, My_SQL

    FillGridWithData
    Dcombos.GetStores Me.DCboStoreName
    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

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
    Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·Œÿ   " & TxtSerial.Text & CHR(13) & " ﬂÊœ   " & TXTCode.Text & CHR(13) & " «·«”„ " & TxtVacName.Text & CHR(13) & "„œÌ— «·Œÿ  " & DcManger.Text & CHR(13) & " „·«ÕŸ«   " & TXTNotes & CHR(13) & " ﬁÌ„… «” Â·«ﬂ «·ÊﬁÊœ ›Ì «·”«⁄Â  " & TXTUsedPowerPriceH & CHR(13) & " ﬁÌ„… «” Â·«ﬂ «·ﬂÂ—»«¡ ›Ì «·”«⁄Â   " & TXTUsedElectricPriceH & CHR(13) & " ﬁÌ„… «” Â·«ﬂ  «·Œÿ «·”«⁄Â   " & TxtLinePriceH & CHR(13) & "«·‘Ì›   " & Me.DcShift.Text
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Line No  " & TxtSerial.Text & CHR(13) & " Code   " & TXTCode.Text & CHR(13) & " Name " & TxtVacName.Text & CHR(13) & "Manger " & DcManger.Text & CHR(13) & " Remarks  " & TXTNotes & CHR(13) & " Power Per Hour  " & TXTUsedPowerPriceH & CHR(13) & " Electric Per Hour   " & TXTUsedElectricPriceH & CHR(13) & " Line Cost Per Hour  " & TxtLinePriceH & CHR(13) & "Shift  " & Me.DcShift.Text
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If
    
End Function

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Me.Caption = "Production Lines"
    Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Line ID"
    Label1(16).Caption = "Depreciation Value"
    Label1(11).Caption = "Line Code"
    Label1(0).Caption = "Line Name"
    Label1(14).Caption = "Line Mngr"
    Label1(5).Caption = "Notes"
    Label1(1).Caption = "Used Power H"
    Label1(4).Caption = "Used Electric H"
    Label1(9).Caption = "Labor Cost H"
    Label1(10).Caption = "Line Cost H"
    Label1(12).Caption = "Shift No"
    Me.TabMain.TabCaption(0) = "Lines's Data"
    Me.TabMain.TabCaption(1) = "Equipments's Data"
    Me.TabMain.TabCaption(2) = "Employee's Data"
    Me.TabMain.TabCaption(3) = "All Lines's Data"
    With Me.Gridx
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("Hourdipp")) = "Depreciation Value"
        .TextMatrix(0, .ColIndex("Ser")) = "I"
        .TextMatrix(0, .ColIndex("name")) = "Equ Name"
        .TextMatrix(0, .ColIndex("UsedPowerPriceH")) = "Used Power Price H"
        .TextMatrix(0, .ColIndex("UsedElectricPriceH")) = "Used Electric Price H"
        .TextMatrix(0, .ColIndex("Notes")) = "Notes"

    End With

    CmdRemove.Caption = "Remove Line"

    With Me.VSFlexGrid1
 
        .TextMatrix(0, .ColIndex("code")) = "code"
 
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("name")) = "Equ Name"
        .TextMatrix(0, .ColIndex("hourprice")) = "Hour Price"

        .TextMatrix(0, .ColIndex("shift1")) = "shift1"
        .TextMatrix(0, .ColIndex("shift2")) = "shift2"
        .TextMatrix(0, .ColIndex("shift3")) = "shift3"
        .TextMatrix(0, .ColIndex("shift4")) = "shift4"
        .TextMatrix(0, .ColIndex("Name1")) = "Swaped Emp"
    End With


   With Me.VSFlexGrid2
 
        .TextMatrix(0, .ColIndex("UserID")) = "code"
 

        .TextMatrix(0, .ColIndex("name")) = "User Name"
    End With
    Label27.Caption = "No Of labors"
    Label29.Caption = "Totals"
    Label5.Caption = "Totals"

    With Me.Grid
 
        .TextMatrix(0, .ColIndex("code")) = "code"
 
        .TextMatrix(0, .ColIndex("Ser")) = "I"
        .TextMatrix(0, .ColIndex("name")) = "Equ Name"
        .TextMatrix(0, .ColIndex("UsedPowerPriceH")) = "Used Power Price H"
        .TextMatrix(0, .ColIndex("UsedElectricPriceH")) = "Used Electric Price H"

    End With

    ALLButton1.Caption = "Remove Line"

    Frame10.Caption = "Labors Work in this line"

    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblProductLine", "id", "")
 
    TxtSerial.Text = StrRecID

    RsSavRec.AddNew
    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Function calnets()
    TXTUsedPowerPriceH = TxtEquiomentPowerTotal

    TXTUsedElectricPriceH = TxtEquiomentElectricTotal
    TxtWorkerPriceH = txt_emp_salary(0)
    TxtHourdipp.Text = TxtHourdippTotal.Text
    TxtLinePriceH = val(TXTUsedPowerPriceH) + val(TXTUsedElectricPriceH) + val(TxtWorkerPriceH)

End Function

Public Sub FiLLRec()
    On Error GoTo ErrTrap
    calnets
    RsSavRec.Fields("name").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)
    RsSavRec.Fields("Code").value = IIf(Me.TXTCode.Text <> "", Trim(TXTCode.Text), Null)

    RsSavRec.Fields("MangerID").value = IIf(IsNumeric(Me.DcManger.BoundText), val(Me.DcManger.BoundText), 0)
    RsSavRec.Fields("ShiftID").value = IIf(IsNumeric(Me.DcShift.BoundText), val(Me.DcShift.BoundText), 0)
    RsSavRec.Fields("HourdippTotal").value = IIf(IsNumeric(Me.TxtHourdippTotal.Text), val(TxtHourdippTotal.Text), 0)
    RsSavRec.Fields("UsedPowerPriceH").value = IIf(IsNumeric(Me.TXTUsedPowerPriceH.Text), val(TXTUsedPowerPriceH.Text), 0)
    RsSavRec.Fields("UsedElectricPriceH").value = IIf(IsNumeric(Me.TXTUsedElectricPriceH.Text), val(TXTUsedElectricPriceH.Text), 0)

    RsSavRec.Fields("WorkerPriceH").value = IIf(IsNumeric(Me.TxtWorkerPriceH.Text), val(TxtWorkerPriceH.Text), 0)
    RsSavRec.Fields("LinePriceH").value = IIf(IsNumeric(Me.TxtLinePriceH.Text), val(TxtLinePriceH.Text), 0)
    RsSavRec.Fields("Notes").value = IIf(Me.TXTNotes.Text <> "", Trim(TXTNotes.Text), Null)
    RsSavRec("StoreID").value = val(Me.DCboStoreName.BoundText)
    RsSavRec.Fields("FormPrint").value = cmbFormPrint.ListIndex + 1
    
    
        If chkIsBasicLine.value = vbChecked Then
            RsSavRec("IsBasicLine").value = 1
        Else
            RsSavRec("IsBasicLine").value = 0
        End If
         
    
    RsSavRec.update

    '   »Ì«‰«  «·⁄„«·
     
    Dim RsEmployee As ADODB.Recordset
    Dim i As Integer
    Dim RsEmpUser As ADODB.Recordset

    If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete From TblProductLineWorker Where LineID=" & val(Me.TxtSerial.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "Delete From TblUsersProductLine Where ProductLineId=" & val(Me.TxtSerial.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    If Me.VSFlexGrid1.Rows <> 1 Then
        Set RsEmployee = New ADODB.Recordset
       ' RsEmployee.Open "TblProductLineWorker", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            
            
            
StrSQL = "SELECT     *  from dbo.TblProductLineWorker Where (1 = -1)"
   RsEmployee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    'Dim RsEmpUser As ADODB.Recordset
        If VSFlexGrid1.Rows > 2 Then
            VSFlexGrid1.Rows = VSFlexGrid1.Rows - 1
        End If
'UserAdmin
        Dim mFound As Boolean
        mFound = True
        For i = 1 To Me.VSFlexGrid1.Rows - 1
            If Trim(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("Emp_id"))) = "" Then
                mFound = False
            Else
                mFound = True
            End If
        Next
        If Not mFound Then
            Dim s As String, mEmp As Long
            s = "SELECT EmpID FROM TblUsers Where UserID = " & UserAdmin
            Set RsEmpUser = New ADODB.Recordset
            RsEmpUser.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not RsEmpUser.EOF Then
                mEmp = val(RsEmpUser!EmpID & "")
                RsEmpUser.Close
                s = "SELECT Emp_Code,Emp_Name  FROM TblEmployee Where Emp_ID = " & mEmp
                RsEmpUser.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Not RsEmpUser.EOF Then
                    s = "INSERT INTO  TblProductLineWorker (EmpID,LineID,EmpCode,EmpIname)" & " VALUES ( " & mEmp & " , " & val(Me.TxtSerial.Text) & ",'" & Trim(RsEmpUser!Emp_Code & "") & "','" & Trim(RsEmpUser!emp_Name & "") & "')"
                    Cn.Execute s
                End If
            End If
        Else
             For i = 1 To Me.VSFlexGrid1.Rows - 1
                RsEmployee.AddNew
                RsEmployee("LineID").value = val(Me.TxtSerial.Text)
                RsEmployee("EmpID").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("Emp_id")))
                RsEmployee("EmpCode").value = VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("code"))
                RsEmployee("EmpIname").value = VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("name"))
                        
                RsEmployee("WorkerPriceH").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("hourprice")))
    
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift1")) = flexChecked Then
                    RsEmployee("shift1").value = 1
                Else
                    RsEmployee("shift1").value = 0
                End If
                
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift2")) = flexChecked Then
                    RsEmployee("shift2").value = 1
                Else
                    RsEmployee("shift2").value = 0
                End If
                
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift3")) = flexChecked Then
                    RsEmployee("shift3").value = 1
                Else
                    RsEmployee("shift3").value = 0
                End If
                
                If VSFlexGrid1.Cell(flexcpChecked, i, VSFlexGrid1.ColIndex("shift4")) = flexChecked Then
                    RsEmployee("shift4").value = 1
                Else
                    RsEmployee("shift4").value = 0
                End If
                        
                RsEmployee.update
            Next i
        End If
       

    End If


If Me.VSFlexGrid2.Rows <> 1 Then
        Set RsEmployee = New ADODB.Recordset
       ' RsEmployee.Open "TblProductLineWorker", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            
            
            
    StrSQL = "SELECT     *  from dbo.TblUsersProductLine Where (1 = -1)"
    RsEmployee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
              For i = 1 To Me.VSFlexGrid2.Rows - 1
                RsEmployee.AddNew
                RsEmployee("ProductLineId").value = val(Me.TxtSerial.Text)
                RsEmployee("userid").value = val(VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("UserID")))
                RsEmployee.update
            Next i
        
       

    End If
  

    '   »Ì«‰«  «·„«ﬂÌ‰« 
     
    Dim RsTblProductLineEquipments As ADODB.Recordset
  
    If Me.TxtModFlg.Text = "E" Then
        StrSQL = "Delete From TblProductLineEquipments Where LineID=" & val(Me.TxtSerial.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    If Me.Gridx.Rows <> 1 Then
        Set RsTblProductLineEquipments = New ADODB.Recordset
      '  RsTblProductLineEquipments.Open "TblProductLineEquipments", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            
                 
StrSQL = "SELECT     *  from dbo.TblProductLineEquipments Where (1 = -1)"
   RsTblProductLineEquipments.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
        If Gridx.Rows > 2 Then
            Gridx.Rows = Gridx.Rows - 1
        End If

        For i = 1 To Me.Gridx.Rows - 1
            RsTblProductLineEquipments.AddNew
            RsTblProductLineEquipments("LineID").value = val(Me.TxtSerial.Text)
            RsTblProductLineEquipments("EquipmentID").value = val(Gridx.TextMatrix(i, Gridx.ColIndex("id")))
            RsTblProductLineEquipments("EquipmentCode").value = Gridx.TextMatrix(i, Gridx.ColIndex("code"))
            RsTblProductLineEquipments("Equipmentname").value = Gridx.TextMatrix(i, Gridx.ColIndex("name"))
            RsTblProductLineEquipments("UsedPowerPriceH").value = val(Gridx.TextMatrix(i, Gridx.ColIndex("UsedPowerPriceH")))
            RsTblProductLineEquipments("Hourdipp").value = val(Gridx.TextMatrix(i, Gridx.ColIndex("Hourdipp")))
            RsTblProductLineEquipments("UsedElectricPriceH").value = val(Gridx.TextMatrix(i, Gridx.ColIndex("UsedElectricPriceH")))
            RsTblProductLineEquipments("Notes").value = Gridx.TextMatrix(i, Gridx.ColIndex("Notes"))
                      
            RsTblProductLineEquipments.update
        Next i

    End If

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    CuurentLogdata
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
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
          
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    VSFlexGrid2.Enabled = True
          
          
    Gridx.Clear flexClearScrollable, flexClearEverything
    Gridx.Rows = 2
    Gridx.Enabled = True
          
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)

    Me.TXTCode.Text = IIf(IsNull(RsSavRec.Fields("code").value), "", RsSavRec.Fields("code").value)
    Me.TXTUsedPowerPriceH.Text = IIf(Not IsNumeric(RsSavRec.Fields("UsedPowerPriceH").value), 0, RsSavRec.Fields("UsedPowerPriceH").value)
    Me.TxtHourdippTotal.Text = IIf(Not IsNumeric(RsSavRec.Fields("HourdippTotal").value), 0, RsSavRec.Fields("HourdippTotal").value)
    Me.TXTUsedElectricPriceH.Text = IIf(Not IsNumeric(RsSavRec.Fields("UsedElectricPriceH").value), 0, RsSavRec.Fields("UsedElectricPriceH").value)
    Me.TxtWorkerPriceH.Text = IIf(Not IsNumeric(RsSavRec.Fields("WorkerPriceH").value), 0, RsSavRec.Fields("WorkerPriceH").value)
    Me.TxtLinePriceH.Text = IIf(Not IsNumeric(RsSavRec.Fields("LinePriceH").value), 0, RsSavRec.Fields("LinePriceH").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(RsSavRec("StoreID").value), "", RsSavRec("StoreID").value)
    If RsSavRec("IsBasicLine").value Then
        chkIsBasicLine.value = vbChecked
    Else
        chkIsBasicLine.value = Unchecked
    End If
    

    If Not IsNull(RsSavRec("MangerID").value) Then
        Me.DcManger.BoundText = IIf(RsSavRec("MangerID").value = 0, 0, RsSavRec("MangerID").value)
    End If

    If Not IsNull(RsSavRec("ShiftID").value) Then
        Me.DcShift.BoundText = IIf(RsSavRec("ShiftID").value = 0, 0, RsSavRec("ShiftID").value)
    End If
 
    Me.cmbFormPrint.ListIndex = IIf(RsSavRec("FormPrint").value = 0, 0, RsSavRec("FormPrint").value - 1)

    Me.TXTNotes.Text = IIf(IsNull(RsSavRec.Fields("Notes").value), "", RsSavRec.Fields("Notes").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                '            Exit Sub
            End If

        Next

    End With

    '»Ì«‰«  «·⁄«„·Ì‰ ›Ì «·Œÿ
    Dim RsEmployee As ADODB.Recordset
    Set RsEmployee = New ADODB.Recordset
    StrSQL = "Select * From TblProductLineWorker Where LineID=" & RsSavRec.Fields("id").value
    StrSQL = StrSQL + " Order By id"
    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then

        With Me.VSFlexGrid1
            .Rows = .FixedRows + RsEmployee.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsEmployee("EmpID").value), 0, val(RsEmployee("EmpID").value))
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(RsEmployee("EmpCode").value), "", RsEmployee("EmpCode").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsEmployee("EmpIname").value), "", RsEmployee("EmpIname").value)
                .TextMatrix(i, .ColIndex("hourprice")) = IIf(IsNull(RsEmployee("WorkerPriceH").value), 0, val(RsEmployee("WorkerPriceH").value))
                .TextMatrix(i, .ColIndex("shift1")) = IIf(IsNull(RsEmployee("Shift1").value), 0, RsEmployee("Shift1").value)
                .TextMatrix(i, .ColIndex("shift2")) = IIf(IsNull(RsEmployee("Shift2").value), 0, RsEmployee("Shift2").value)
                .TextMatrix(i, .ColIndex("shift3")) = IIf(IsNull(RsEmployee("Shift3").value), 0, RsEmployee("Shift3").value)
                .TextMatrix(i, .ColIndex("shift4")) = IIf(IsNull(RsEmployee("Shift4").value), 0, RsEmployee("Shift4").value)
                       
                RsEmployee.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
                    
        End With

    End If
        
        
  Dim RsUsers As ADODB.Recordset
    Set RsUsers = New ADODB.Recordset
    StrSQL = "Select tblUsers.UserName,TblUsersProductLine.* From TblUsersProductLine "
    StrSQL = StrSQL & " Inner join tblUsers On tblUsers.UserId =TblUsersProductLine.UserId "
    StrSQL = StrSQL & " Where ProductLineId=" & RsSavRec.Fields("id").value
    StrSQL = StrSQL + " Order By id"
    RsUsers.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUsers.BOF Or RsUsers.EOF) Then

        With Me.VSFlexGrid2
            .Rows = .FixedRows + RsUsers.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("UserID")) = IIf(IsNull(RsUsers("UserID").value), 0, val(RsUsers("UserID").value))

                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsUsers("UserName").value), "", RsUsers("UserName").value)
                       
                RsUsers.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
                    
        End With

    End If
        
    '»Ì«‰«  «·„«ﬂÌ‰«   ›Ì «·Œÿ
    Dim RsTblProductLineEquipments As ADODB.Recordset
    Set RsTblProductLineEquipments = New ADODB.Recordset
    StrSQL = " SELECT     dbo.TblProductLineEquipments.id, dbo.TblProductLineEquipments.Equipmentname, dbo.TblProductLineEquipments.UsedPowerPriceH,"
    StrSQL = StrSQL + "                  dbo.TblProductLineEquipments.UsedElectricPriceH, dbo.TblProductLineEquipments.Notes, dbo.TblProductLineEquipments.Hourdipp,"
    StrSQL = StrSQL + "                  dbo.TblProductLineEquipments.EquipmentCode, dbo.TblProductLineEquipments.LineID, dbo.TblProductLineEquipments.EquipmentID, dbo.TblEquipments.name,"
    StrSQL = StrSQL + "                  dbo.TblEquipments.code , dbo.TblEquipments.NameE"
    StrSQL = StrSQL + "  FROM         dbo.TblProductLineEquipments LEFT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblEquipments ON dbo.TblProductLineEquipments.EquipmentID = dbo.TblEquipments.id"
    StrSQL = StrSQL + " Where dbo.TblProductLineEquipments.LineID =" & RsSavRec.Fields("id").value
    StrSQL = StrSQL + " Order By dbo.TblProductLineEquipments.id "
    RsTblProductLineEquipments.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTblProductLineEquipments.BOF Or RsTblProductLineEquipments.EOF) Then

        With Me.Gridx
            .Rows = .FixedRows + RsTblProductLineEquipments.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsTblProductLineEquipments("EquipmentID").value), 0, val(RsTblProductLineEquipments("EquipmentID").value))
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(RsTblProductLineEquipments("Code").value), "", RsTblProductLineEquipments("Code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsTblProductLineEquipments("name").value), "", RsTblProductLineEquipments("name").value)
                Else
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsTblProductLineEquipments("NameE").value), "", RsTblProductLineEquipments("NameE").value)
                End If
                .TextMatrix(i, .ColIndex("UsedPowerPriceH")) = IIf(IsNull(RsTblProductLineEquipments("UsedPowerPriceH").value), 0, val(RsTblProductLineEquipments("UsedPowerPriceH").value))
                .TextMatrix(i, .ColIndex("UsedElectricPriceH")) = IIf(IsNull(RsTblProductLineEquipments("UsedElectricPriceH").value), 0, val(RsTblProductLineEquipments("UsedElectricPriceH").value))
                .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(RsTblProductLineEquipments("Notes").value), "", RsTblProductLineEquipments("Notes").value)
                .TextMatrix(i, .ColIndex("Hourdipp")) = IIf(IsNull(RsTblProductLineEquipments("Hourdipp").value), 0, val(RsTblProductLineEquipments("Hourdipp").value))
                RsTblProductLineEquipments.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
                    
        End With

    End If
        
    ReLineGrid
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
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Private Sub Gridx_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Gridx

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
                StrSQL = "SELECT  *   from TblEquipments Where id= " & val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
                        .TextMatrix(Row, .ColIndex("Hourdipp")) = IIf(IsNull(rs("Hourdipp").value), 0, rs("Hourdipp").value)
                        .TextMatrix(Row, .ColIndex("UsedPowerPriceH")) = IIf(IsNull(rs("UsedPowerPriceH").value), "", rs("UsedPowerPriceH").value)
                        .TextMatrix(Row, .ColIndex("UsedElectricPriceH")) = IIf(IsNull(rs("UsedElectricPriceH").value), 0, rs("UsedElectricPriceH").value)
                        .TextMatrix(Row, .ColIndex("Notes")) = IIf(IsNull(rs("Notes").value), "", rs("Notes").value)
                            
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "code"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from TblEmployee Where Emp_Code=" & .TextMatrix(Row, Col)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
          
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                    .TextMatrix(Row, .ColIndex("id")) = ""
                    .TextMatrix(Row, .ColIndex("name")) = ""
                End If

        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

End Sub

Private Sub Gridx_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With Gridx
 
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub

            Case "UsedPowerPriceH"
                Cancel = True

            Case "UsedElectricPriceH"
                Cancel = True
            Case "Hourdipp"
                Cancel = True
        End Select

    End With

    Gridx.ComboList = ""
End Sub

Private Sub Gridx_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
    With Gridx

        Select Case .ColKey(Col)

            Case "name"
                StrSQL = "select * from TblEquipments"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = VSFlexGrid1.BuildComboList(rs, "name", "id")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap

    If RecId = 0 Then RecId = 1
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1

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
        '   btnNext.Enabled = False
        '   btnPrevious.Enabled = False
        '   btnFirst.Enabled = False
        '   btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.Text <> "" Then
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
    My_SQL = "select * From TblProductLine order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
           
                .TextMatrix(i, .ColIndex("UsedPowerPriceH")) = IIf(IsNull(rs.Fields("UsedPowerPriceH").value), "", rs.Fields("UsedPowerPriceH").value)
            
                .TextMatrix(i, .ColIndex("UsedElectricPriceH")) = IIf(IsNull(rs.Fields("UsedElectricPriceH").value), "", rs.Fields("UsedElectricPriceH").value)
            
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs.Fields("Code").value), "", rs.Fields("Code").value)
            
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

Private Function CheckDelCountry(Lngid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
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



Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
                                  
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
    
    'ddd

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "name"

        
                StrSQL = "select * from TblUsers"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = VSFlexGrid2.BuildComboList(rs, "UserName", "UserID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With






    Dim StrAccountCode As String
    
    
    
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim EmployeeSalary As Double
    With VSFlexGrid2

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
'                StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/208,2) as hourprice,* from TblEmployee Where Emp_id= " & val(StrAccountCode)
StrSQL = "SELECT      * from TblUsers Where UserID= " & val(StrAccountCode)

  
 
                        
                          
      
                    
                    
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("UserID")) = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
                        
                        '.TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                        '.TextMatrix(Row, .ColIndex("hourprice")) = IIf(IsNull(rs("hourprice").value), "", rs("hourprice").value)
                       
                            
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
 
          
        
            Case "UserID"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from TblUsers Where UserID=" & val(.TextMatrix(Row, Col))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
          
                 '   .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                    .TextMatrix(Row, .ColIndex("UserID")) = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
                    
                Else
                    .TextMatrix(Row, .ColIndex("id")) = ""
                    .TextMatrix(Row, .ColIndex("name")) = ""
                    .TextMatrix(Row, .ColIndex("UserID")) = ""
                End If

        
 
  
                        
                          
      
                    
                    
         
        
        
        
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    
End Sub


Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim EmployeeSalary As Double
    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
'                StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/208,2) as hourprice,* from TblEmployee Where Emp_id= " & val(StrAccountCode)
StrSQL = "SELECT      * from TblEmployee Where Emp_id= " & val(StrAccountCode)

  EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(StrAccountCode), "")
 
                        
                          
      
                    
                    
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("FULLCODE").value), "", rs("FULLCODE").value)
                        .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                        '.TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                        '.TextMatrix(Row, .ColIndex("hourprice")) = IIf(IsNull(rs("hourprice").value), "", rs("hourprice").value)
                            .TextMatrix(Row, .ColIndex("hourprice")) = Round(EmployeeSalary / 240, 2)
                            
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
 
            Case "Name1"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
                StrSQL = "SELECT        * from TblEmployee Where Emp_id= " & val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                   
                        .TextMatrix(Row, .ColIndex("Emp_id1")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                       ' .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
        
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "code"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from TblEmployee Where FULLCODE=" & .TextMatrix(Row, Col)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
          
                 '   .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                    .TextMatrix(Row, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                    
                Else
                    .TextMatrix(Row, .ColIndex("id")) = ""
                    .TextMatrix(Row, .ColIndex("name")) = ""
                    .TextMatrix(Row, .ColIndex("Emp_id")) = ""
                End If

        
 
  EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(Row, .ColIndex("Emp_id"))), "")
 .TextMatrix(Row, .ColIndex("hourprice")) = Round(EmployeeSalary / 240, 2)
                        
                          
      
                    
                    
         
        
        
        
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid1.ComboList = ""

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "name"

        
                StrSQL = "select * from TblEmployee"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = VSFlexGrid1.BuildComboList(rs, "Emp_Name", "Emp_ID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub





Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid2.ComboList = ""

End Sub





Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "name"

        
                StrSQL = "select * from TblUsers"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = VSFlexGrid2.BuildComboList(rs, "UserName", "UserID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub
