VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmIqarCompnent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»„ﬂÊ‰«  «·⁄ﬁ«—« "
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   Icon            =   "FrmIqarCompnent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   30
      Width           =   7785
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   28
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   -105
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
            TabIndex        =   29
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   510
         Visible         =   0   'False
         Width           =   945
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
               Picture         =   "FrmIqarCompnent.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarCompnent.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   30
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
         ButtonImage     =   "FrmIqarCompnent.frx":245A
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   31
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
         ButtonImage     =   "FrmIqarCompnent.frx":27F4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   32
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
         ButtonImage     =   "FrmIqarCompnent.frx":2B8E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   33
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
         ButtonImage     =   "FrmIqarCompnent.frx":2F28
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2400
         Picture         =   "FrmIqarCompnent.frx":32C2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬂÊ‰«  «·⁄ﬁ«—« "
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
         Left            =   5325
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   90
         Width           =   2280
      End
   End
   Begin VB.TextBox txtcolor 
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
      Left            =   8280
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   720
      Width           =   1065
   End
   Begin VB.TextBox Text1 
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
      Left            =   8280
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   120
      Width           =   1065
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   555
      Width           =   7755
      Begin VB.TextBox TxtNameE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txtid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   495
         Width           =   5895
      End
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "FrmIqarCompnent.frx":6F2A
         Left            =   2280
         List            =   "FrmIqarCompnent.frx":6F3A
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4950
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo DcAccountsus 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
         Top             =   1200
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XPCboGroup 
         Height          =   315
         Left            =   2040
         TabIndex        =   44
         Top             =   1560
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„Ã„Ê⁄…"
         Height          =   285
         Index           =   4
         Left            =   5970
         TabIndex        =   45
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ Õ”«» «·„’—Ê›"
         Height          =   285
         Index           =   3
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1320
         Width           =   1890
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„Ã„Ê⁄Â ≈‰Ã·Ì“Ì"
         Height          =   285
         Index           =   2
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   855
         Width           =   1890
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„ÃÊ⁄Â"
         Height          =   285
         Index           =   0
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   120
         Width           =   1890
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„Ã„Ê⁄Â ⁄—»Ì"
         Height          =   285
         Index           =   1
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   510
         Width           =   1890
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6240
      Width           =   7680
      _cx             =   13547
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
         Left            =   5325
         TabIndex        =   1
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
         ButtonImage     =   "FrmIqarCompnent.frx":6F53
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   3780
         TabIndex        =   2
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
         ButtonImage     =   "FrmIqarCompnent.frx":72ED
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   4545
         TabIndex        =   3
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
         ButtonImage     =   "FrmIqarCompnent.frx":7687
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   3000
         TabIndex        =   4
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
         ButtonImage     =   "FrmIqarCompnent.frx":7A21
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   2250
         TabIndex        =   5
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
         ButtonImage     =   "FrmIqarCompnent.frx":7DBB
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   1530
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
         ButtonImage     =   "FrmIqarCompnent.frx":8355
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   7
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
         ButtonImage     =   "FrmIqarCompnent.frx":86EF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   4725
         TabIndex        =   8
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
         ButtonImage     =   "FrmIqarCompnent.frx":8A89
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   1455
         TabIndex        =   9
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
         ButtonImage     =   "FrmIqarCompnent.frx":8E23
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin XtremeSuiteControls.CommonDialog CommonDialog1 
         Left            =   5040
         Top             =   1320
         _Version        =   786432
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   4
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   1410
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   975
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   900
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   3780
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7740
      _cx             =   13653
      _cy             =   6668
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
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   3300
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   7755
         _cx             =   13679
         _cy             =   5821
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmIqarCompnent.frx":91BD
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   21
         Left            =   6960
         TabIndex        =   21
         Top             =   3720
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   476
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–›"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmIqarCompnent.frx":930D
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   0
         Left            =   6840
         TabIndex        =   41
         Top             =   3360
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   476
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
         ButtonImage     =   "FrmIqarCompnent.frx":98A7
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   675
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
End
Attribute VB_Name = "FrmIqarCompnent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic

Dim RecId As String
Dim II As Long
Public LngRow As Long
Public StrAccountCodepu As String
Public LngCol As Long
Public idvac As Long
Public empid As Long
Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.Text <> "" Then
    

        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
            RsSavRec.Find "ID=" & val(TxtVac_ID.Text), , adSearchForward, 1
            'CuurentLogdata ("D")
            RsSavRec.delete
            StrSQL = "Delete From TblAqrCompenetDet Where IDAqComp=" & val(Me.TxtVac_ID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

       ' StrSQL = "Delete From Supervisors Where DeparmentID=" & val(Me.TxtVac_ID.text)
      '      Cn.Execute StrSQL, , adExecuteNoRecords
              fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 2
            fg.Enabled = True
         
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
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
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
          fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 2
            fg.Enabled = True
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
          fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 2
            fg.Enabled = True
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
fg.Rows = fg.Rows + 1
            fg.Enabled = True
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        'FiLLRec
      '  Me.TxtVacName.SetFocus
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
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    fg.Clear flexClearScrollable, flexClearEverything
    
fg.Rows = 2
            fg.Enabled = True
          
    TxtModFlg.Text = "N"
'txtid.text = TxtVac_ID.text
    My_SQL = "TblAqrCompenet"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtVac_ID.Text = rs.RecordCount + 1
    Else
        TxtVac_ID.Text = 1
    End If

    rs.Close
  
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
          fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 2
            fg.Enabled = True
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
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
          fg.Clear flexClearScrollable, flexClearEverything
            fg.Rows = 2
            fg.Enabled = True
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
   Dim StrSQL As String
    '---------------------- check if data Vaclete -----------------------
 If Me.TxtName.Text = "" Then
            Msg = "ÌÃ» ﬂ «»… «”„ «·„Ã„Ê⁄Â   !! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtName.SetFocus
       '     SendKeys "{F4}"
            Exit Sub
        End If
  
    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

   StrVacName = IsRecExist("TblAqrCompenet", "Name", Trim(TxtName.Text), "Name", "ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
       Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·„Ã„Ê⁄Â „‰ ﬁ»·"
   
       MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
       TxtName.SetFocus
    
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
StrSQL = "Delete From TblAqrCompenetDet Where IDAqComp=" & val(Me.TxtVac_ID.Text)
         Cn.Execute StrSQL, , adExecuteNoRecords

        'StrSQL = "Delete From Supervisors Where DeparmentID=" & val(Me.TxtVac_ID.text)
        '   Cn.Execute StrSQL, , adExecuteNoRecords    '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub
 
Private Sub BtnUndo_Click()
BtnFirst_Click
    'FindRec val(txtid.text)
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
Cmd(21).Caption = "Delete"
    Me.Caption = "IqarCompnent"
lbl(0).Caption = "Op No"
lbl(1).Caption = "Group Arbic"
lbl(2).Caption = "Group Eng"
    Label1(2).Caption = Me.Caption

     With Me.fg
     .TextMatrix(0, .ColIndex("name")) = "Name Arbic"
     .TextMatrix(0, .ColIndex("namee")) = "Name Eng"
        .TextMatrix(0, .ColIndex("serial")) = "Serial"
        .TextMatrix(0, .ColIndex("price")) = " Price"
       ' .TextMatrix(0, .ColIndex("Emp_name")) = "Tchnicians Name "

'.TextMatrix(0, .ColIndex("audStr1")) = "No.Tchnicians"
    End With


 
    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
       
    IntCounter = 0

    With fg

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
   
            End If

        Next i
 
    End With
    
End Sub



Private Sub Cmd_Click(Index As Integer)

Select Case Index
Case 0
    With Me.fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Select
    ReLineGrid

End Sub











Private Sub DcAccountsus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 20150204
    End If
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String

    With fg
               
    

     '   Select Case .ColKey(Col)
 
     '       Case "Emp_name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
     '           StrAccountCode = .ComboData
     '           LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID1"), False, True)
     '           .TextMatrix(Row, .ColIndex("EmpID1")) = StrAccountCode
                'StrAccountCodepu = StrAccountCode
               ' StrSQL = "select * from TblExtraExpeneses where Id=" & val(StrAccountCode)
               ' rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               '  If rs.RecordCount > 0 Then
               '     .TextMatrix(Row, .ColIndex("typeexpen")) = IIf(IsNull(rs("TypeExtrExpen").value), 0, rs("TypeExtrExpen").value)
              '  Else
               '     .TextMatrix(Row, .ColIndex("typeexpen")) = ""
               ' End If
     '              End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

'Private Sub Fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'With fg

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
'        Select Case .ColKey(Col)
            
'            Case "EmpID"
'               Cancel = True
          
               '  Case "comp"
               ' fg.ComboList = ""
               '  Case "bill"
               ' fg.ComboList = ""
'        End Select

'    fg.ComboList = ""
'    End With
'End Sub

'Private Sub Fg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
' Dim LngItemID As Long
'    Dim LngStoreID As Long
'    Dim rdate As Date
'  ' Dim frm As FrmGridAddItemComment
   ' Dim Frm1 As ItemProductionDate

    'On Error GoTo ErrTrap

 '   With Me.fg

 '       Select Case .ColKey(Col)

 '                Case "audStr1"
                 ' If .TextMatrix(Row, .ColIndex("Emp_name")) = "" Then
             '  MsgBox "ÌÃ» «Œ Ì«— «·„‘—› «Ê·«"
              ' Exit Sub
 '             'Else
 '                 LngRow = Row
 'LngCol = Col
 ' idvac = val(TxtVac_ID.text)
  
 'EmpID = fg.TextMatrix(Row, .ColIndex("EmpID1"))
 
             ' ItemProductionDate Row, Col, , 1
               'Load FromTechnicians
               'FromTechnicians.show
              ' End If
           '       Case "timEnter"
                                    
'
' LngCol = Col
'                  Load ItemProductionDate
'                ItemProductionDate.show
             'ItemProductionDate Row, Col, , 1
           '     End Select
           '     End With

'End Sub

'Private Sub Fg_Click()
'fisbtixt
'End Sub

'Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'Dim rs As New ADODB.Recordset
'    Dim StrSQL  As String
'    Dim StrAccountType As String
'    Dim StrComboList As String
'    Dim Msg As String
'
'
'    With fg

'        Select Case .ColKey(Col)

'            Case "Emp_name"
       
' StrSQL = " select * from  TblEmployee "
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrComboList = fg.BuildComboList(rs, "Emp_Name", "Emp_ID")
'                Else
'                    StrComboList = fg.BuildComboList(rs, "Emp_Namee", "Emp_ID")
'                End If
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'
'                .ComboList = StrComboList
'                 Case "audStr"
'                .ColComboList(.ColIndex("audStr")) = "..."
'        End Select
'
'    End With
'End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
' Dim Dcombos As ClsDataCombos
' Set Dcombos = New ClsDataCombos
'    Dcombos.GetEmpDepartmentCar Me.DcbDept
'    Dcombos.GetEmployees Me.DcboEmpName
  Dim Dcombos As ClsDataCombos
   Set Dcombos = New ClsDataCombos
 Dcombos.GetAccountingCodes Me.DcAccountsus, True

    
    Dcombos.GetItemSGroups Me.XPCboGroup

    My_SQL = "TblAqrCompenet"
    Set BKGrndPic = New ClsBackGroundPic
   Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
 
    Me.TxtModFlg.Text = "R"

   ScreenNameArabic = "  „ﬂÊ‰«  «·⁄ﬁ«—«  "
    ScreenNameEnglish = " Iqar Compenet"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Resize_Form Me
'
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
'
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
   fill_combo DCUser, My_SQL

  '

  '  With Me.Grid
  '      .Cell(flexcpPicture, 0, .ColIndex("DepartmentName")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
   '     .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
'
'        For i = 0 To .Cols - 1
'            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
'        Next
'
'       .ExtendLastCol = True
'       .WallPaper = BKGrndPic.Picture
'        .RowHeight(-1) = 300
'    End With

    BtnFirst_Click
    ShowTip
FillGridWithData
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

'Function CuurentLogdata(Optional Currentmode As String)
   '  LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & "ﬂÊœ " & TxtSerial.text & Chr(13) & "   «”„ «·ﬁ”„ " & TxtVacName
   ''     LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & "Code  " & TxtSerial.text & Chr(13) & "   Name " & TxtVacNamee
    '   If Currentmode <> "D" Then
   '     AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
   ' Else
   '     AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
   ' End If
    
'End Function

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblAqrCompenet", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    TxtVac_ID.Text = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
Dim i, j As Integer
Dim test_split() As String
Dim s As String
Dim test_split1() As String
Dim s1 As String
Dim Sql As String
Dim nElements As Integer
    On Error GoTo ErrTrap
    Dim StrSQL  As String
Dim RsDetails1 As ADODB.Recordset
Dim RsDetails As ADODB.Recordset


            
    RsSavRec.Fields("ID").value = val(IIf(TxtVac_ID.Text <> "", Me.TxtVac_ID.Text, 0))
    RsSavRec.Fields("Name").value = IIf(Me.TxtName.Text <> "", Me.TxtName.Text, Null)
    RsSavRec.Fields("Namee").value = IIf(Me.TxtNameE.Text <> "", Me.TxtNameE.Text, Null)
    RsSavRec.Fields("Accountsus").value = IIf(DcAccountsus.BoundText <> "", (DcAccountsus.BoundText), Null)
    
    RsSavRec.Fields("GroupId").value = IIf(XPCboGroup.BoundText <> "", (XPCboGroup.BoundText), Null)
    
     '     sql = "update TblEmployee set   WorkShop_Job=1  where Emp_ID=" & val(Me.DcboEmpName.BoundText) & ""
  
     '                               Cn.Execute sql
    ' RsSavRec.Fields("DeptColor").value = IIf(txtcolor.text <> "", Trim(txtcolor.text), Null)
     RsSavRec.update
     If XPCboGroup.Text <> "" Then
        s = "Update Groups Set AqrCompenetId = " & val(txtid) & " Where GroupId = " & val(XPCboGroup.BoundText)
        Cn.Execute s
    End If
         Set RsDetails1 = New ADODB.Recordset
       RsDetails1.Open "TblAqrCompenetDet", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If fg.Rows > 1 Then
                ' fg2.Rows = fg2.Rows - 1
          
       For i = Me.fg.FixedRows To fg.Rows - 2
       If fg.TextMatrix(i, fg.ColIndex("name")) <> "" Then
           RsDetails1.AddNew
          RsDetails1("IDAqComp").value = val(TxtVac_ID.Text) 'val(IIf(TxtVac_ID.text <> "", TxtVac_ID.text, 0))
RsDetails1("Accountsus").value = DcAccountsus.BoundText

            RsDetails1("Price").value = val(fg.TextMatrix(i, fg.ColIndex("price")))
         RsDetails1("Name").value = IIf(fg.TextMatrix(i, fg.ColIndex("name")) <> "", fg.TextMatrix(i, fg.ColIndex("name")), Null)
          RsDetails1("Namee").value = IIf(fg.TextMatrix(i, fg.ColIndex("namee")) <> "", fg.TextMatrix(i, fg.ColIndex("namee")), Null)
    '   sql = "update TblEmployee set   WorkShop_Job=2  where Emp_ID=" & val(fg.TextMatrix(i, fg.ColIndex("EmpID1"))) & ""

    '                                Cn.Execute sql
    RsDetails1.update
           End If
         
     
    
           Next i
        End If

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'  FillGridWithData
   FiLLTXT
    TxtModFlg = "R"

 'CuurentLogdata
    Exit Sub
    
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
         
 '   TxtModFlg = "R"
'    FillGridWithData
 '     MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
FiLLTXT
End Sub

Public Sub FiLLTXT(Optional Lngid As Integer = 0)
    On Error GoTo ErrTrap
   Dim RsDetails1 As ADODB.Recordset
    Dim RsDetails As ADODB.Recordset
 Dim StrSQL As String
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtName.Text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
        TxtNameE.Text = IIf(IsNull(RsSavRec.Fields("Namee").value), "", RsSavRec.Fields("Namee").value)
          Me.DcAccountsus.BoundText = IIf(IsNull(RsSavRec.Fields("Accountsus").value), "", RsSavRec.Fields("Accountsus").value)
          
          
          
          
    XPCboGroup.BoundText = IIf(IsNull(RsSavRec.Fields("GroupId").value), "", RsSavRec.Fields("GroupId").value)

   'Me.DcbDept.BoundText = IIf(IsNull(RsSavRec.Fields("DeparmentID").value), "", RsSavRec.Fields("DeparmentID").value)
   txtid.Text = TxtVac_ID.Text
      
     

'Set RsDetails = New ADODB.Recordset
'StrSQL = "SELECT     dbo.TblAqrCompenet.ID, dbo.TblAqrCompenet.Name, dbo.TblAqrCompenetDet.IDAqComp, dbo.TblAqrCompenetDet.Name AS NameDet, "
'  StrSQL = StrSQL & "                    dbo.TblAqrCompenetDet.Price"
'StrSQL = StrSQL & " FROM         dbo.TblAqrCompenet LEFT OUTER JOIN"
' StrSQL = StrSQL & "                     dbo.TblAqrCompenetDet ON dbo.TblAqrCompenet.ID = dbo.TblAqrCompenetDet.ID"
'StrSQL = StrSQL & " Where (dbo.TblAqrCompenetDet.IDAqComp = " & val(TxtVac_ID.text)
''StrSQL = StrSQL & " Where (dbo.Technicians1.DeparmentID = " & val(TxtVac_ID.text) & ")"
'' StrSQL = StrSQL & " Where (dbo.Technicians.DeparmentID =" & val(TxtVac_ID.text) & ") and (dbo.Technicians.Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID1"))) & ") "

'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'     If Not (RsDetails.BOF Or RsDetails.EOF) Then
'       With fg
'        RsDetails.MoveFirst
'       .Rows = .FixedRows + RsDetails.RecordCount
'
'        For i = .FixedRows To .Rows - 1
'
'            .TextMatrix(i, .ColIndex("serial")) = i
'            .TextMatrix(i, .ColIndex("ID")) = val(IIf(IsNull(RsDetails("ID").value), 0, RsDetails("ID").value))
'            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("Name").value), "", RsDetails("Name").value) 'RsDetails1("Value").value
'.TextMatrix(i, .ColIndex("price")) = val(IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value))
'
 
                   ' If SystemOptions.UserInterface = ArabicInterface Then
                   ' .TextMatrix(i, .ColIndex("Emp_name")) = IIf(IsNull(RsDetails("Emp_Name").value), "", RsDetails("Emp_Name").value)
               ' Else
               '    .TextMatrix(i, .ColIndex("Emp_name")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value)
            ' End If
'            RsDetails.MoveNext
 
'        Next i
'End With

'    End If
' ReLineGrid
'  Next j
  
  FillGridWithData
 ' ReLineGrid
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
RsDetails1.Close
 Set RsDetails1 = Nothing
'RsDetails1.Close
' Set RsDetails1 = Nothing
  '  fillapprovData
  

 

   ' With fg
'
'        For i = 1 To .Rows - 1
'
'            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("DeparmentID")) Then
''                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
 '               .Row = i
 '               Exit Sub
 '           End If

 '       Next

   ' End With

ErrTrap:

End Sub







Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

'Private Sub TxtSearchCode1_KeyPress(KeyAscii As Integer)
'Dim EmpID As Integer

'    If KeyAscii = vbKeyReturn Then
'        GetEmployeeIDFromCode TxtSearchCode1.text, EmpID
'        DcboEmpName.BoundText = EmpID
'    End If
'End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap

    RsSavRec.Find "IDAqComp=" & RecId, , adSearchForward, 1

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
        fg.Enabled = True
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '   btnNext.Enabled = False
        '   btnPrevious.Enabled = False
        '   btnFirst.Enabled = False
        '   btnLast.Enabled = False

    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        fg.Enabled = True
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
        fg.Enabled = True
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
    My_SQL = "SELECT     dbo.TblAqrCompenet.ID, dbo.TblAqrCompenet.Name, dbo.TblAqrCompenetDet.IDAqComp, dbo.TblAqrCompenetDet.Name AS NameDet, "
  My_SQL = My_SQL & "                    dbo.TblAqrCompenetDet.Price , dbo.TblAqrCompenetDet.Namee AS NameEDet "
My_SQL = My_SQL & " FROM         dbo.TblAqrCompenet LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAqrCompenetDet ON dbo.TblAqrCompenet.ID = dbo.TblAqrCompenetDet.IDAqComp"
My_SQL = My_SQL & "  WHERE     (dbo.TblAqrCompenetDet.IDAqComp = " & val(TxtVac_ID.Text) & ")"
 '   rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.fg
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("serial")) = i
                
              '  Me.txtcolor.text = IIf(IsNull(rs.Fields("DeptColor").value), "", rs.Fields("DeptColor").value)
              ' .Cell(flexcpBackColor, i, 1, i, 3) = Me.txtcolor.text  ''''''IIf(IsNull(rs.Fields("DeptColor").value), 47777, rs.Fields("DeptColor").value)
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("NameEDet").value), "", rs.Fields("NameEDet").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("NameDet").value), "", rs.Fields("NameDet").value)

                .TextMatrix(i, .ColIndex("price")) = val(IIf(IsNull(rs.Fields("price").value), 0, rs.Fields("price").value))
            
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

'Private Function CheckDelDepartment(LngDepartmentID As Long) As Boolean
'    Dim rs As ADODB.Recordset
'    Dim StrSQL As String
'    StrSQL = "Select * From TblEmployee Where DepartmentID=" & LngDepartmentID & ""
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        CheckDelDepartment = False
'    Else
'        CheckDelDepartment = True
'    End If

'    rs.Close
'    Set rs = Nothing
'End Function
