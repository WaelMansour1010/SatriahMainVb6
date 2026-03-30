VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmOldContract 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13485
   Icon            =   "FrmOldContract.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»«⁄Â «·⁄ﬁÊœ/«·›Ê« Ì— «·„‰ ÂÌ…"
      Height          =   1020
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6252
      Width           =   5775
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   3540
         TabIndex        =   26
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   102760449
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   1260
         TabIndex        =   27
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   102760449
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton BtnPrint11 
         Height          =   168
         Left            =   2160
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1488
         _ExtentX        =   2619
         _ExtentY        =   291
         ButtonStyle     =   1
         ButtonPositionImage=   2
         Caption         =   "ÿ»«⁄Â"
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
         ButtonImage     =   "FrmOldContract.frx":57E2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   510
         Left            =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   900
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… "
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
         ButtonImage     =   "FrmOldContract.frx":5B7C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   1
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   0
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -90
      Width           =   13515
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   4
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
            TabIndex        =   5
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
         TabIndex        =   2
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
         TabIndex        =   1
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
               Picture         =   "FrmOldContract.frx":C3DE
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":C778
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":CB12
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":CEAC
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":D246
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":D5E0
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":D97A
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmOldContract.frx":DF14
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   150
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
         ButtonImage     =   "FrmOldContract.frx":E2AE
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
         Top             =   150
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
         ButtonImage     =   "FrmOldContract.frx":E648
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   150
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
         ButtonImage     =   "FrmOldContract.frx":E9E2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   150
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
         ButtonImage     =   "FrmOldContract.frx":ED7C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "  »Ì«‰«  «·⁄ﬁÊœ/Ê›Ê« Ì— «·„»Ì⁄«  «·ﬁœÌ„…  "
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
         Left            =   6135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   7230
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   6000
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6252
      Width           =   7440
      _cx             =   13123
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
         Left            =   5175
         TabIndex        =   12
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
         ButtonImage     =   "FrmOldContract.frx":F116
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   3360
         TabIndex        =   13
         Top             =   600
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
         ButtonImage     =   "FrmOldContract.frx":F4B0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   4275
         TabIndex        =   14
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
         ButtonImage     =   "FrmOldContract.frx":F84A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   2505
         TabIndex        =   15
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
         ButtonImage     =   "FrmOldContract.frx":FBE4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   1500
         TabIndex        =   16
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
         ButtonImage     =   "FrmOldContract.frx":FF7E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   6030
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   540
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
         ButtonImage     =   "FrmOldContract.frx":10518
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   18
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
         ButtonImage     =   "FrmOldContract.frx":108B2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   705
         TabIndex        =   19
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
         ButtonImage     =   "FrmOldContract.frx":10C4C
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
         TabIndex        =   23
         Top             =   225
         Width           =   540
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   2505
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   975
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   2925
      Left            =   0
      TabIndex        =   24
      Top             =   600
      Width           =   13485
      _cx             =   23786
      _cy             =   5159
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmOldContract.frx":10FE6
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
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   2610
      Left            =   0
      TabIndex        =   32
      Top             =   3600
      Width           =   13485
      _cx             =   23786
      _cy             =   4604
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·»Ì«‰«  «·«”«”Ì… |»Ì«‰«  „ Œ’’…"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
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
      Begin VB.Frame Frm2 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2190
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   45
         Width           =   13395
         Begin VB.TextBox Txt_ReturnValue 
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
            Left            =   10080
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ﬁÌ„… «·⁄ﬁœ"
            Top             =   707
            Visible         =   0   'False
            Width           =   828
         End
         Begin VB.TextBox Txt_ReturnNo 
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
            Left            =   11400
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· —ﬁ„ «·⁄ﬁœ"
            Top             =   707
            Visible         =   0   'False
            Width           =   948
         End
         Begin VB.TextBox TxtContractNo 
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
            Left            =   7140
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· —ﬁ„ «·”‰œ"
            Top             =   707
            Width           =   1425
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
            Left            =   11700
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   276
            Width           =   1308
         End
         Begin VB.ComboBox CmbType 
            BackColor       =   &H80000018&
            Height          =   315
            ItemData        =   "FrmOldContract.frx":111E9
            Left            =   2280
            List            =   "FrmOldContract.frx":111F9
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   2310
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TXtContractValue 
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
            TabIndex        =   46
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· «·ﬁÌ„…  "
            Top             =   707
            Width           =   828
         End
         Begin VB.TextBox TxtRemarks 
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
            Height          =   435
            Left            =   120
            MaxLength       =   50
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   1680
            Width           =   4536
         End
         Begin VB.TextBox TxtValue 
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
            Left            =   7200
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1680
            Width           =   1530
         End
         Begin VB.ComboBox DcbPaymentType 
            Height          =   288
            Left            =   10140
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1680
            Width           =   1785
         End
         Begin VB.TextBox TxtNetValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
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
            Left            =   5550
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1680
            Width           =   1530
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   0
            Width           =   7605
            Begin XtremeSuiteControls.RadioButton RdType 
               Height          =   252
               Index           =   0
               Left            =   4680
               TabIndex        =   39
               Top             =   240
               Width           =   852
               _Version        =   786432
               _ExtentX        =   1503
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "⁄ﬁÊœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdType 
               Height          =   252
               Index           =   1
               Left            =   3240
               TabIndex        =   40
               Top             =   240
               Width           =   852
               _Version        =   786432
               _ExtentX        =   1503
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "›Ê« Ì—"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdType 
               Height          =   252
               Index           =   2
               Left            =   1800
               TabIndex        =   99
               Top             =   240
               Width           =   852
               _Version        =   786432
               _ExtentX        =   1503
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "„— Ã⁄« "
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdType 
               Height          =   252
               Index           =   3
               Left            =   240
               TabIndex        =   104
               Top             =   240
               Width           =   852
               _Version        =   786432
               _ExtentX        =   1503
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "«·”‰œ« "
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
               Height          =   285
               Index           =   10
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   240
               Width           =   930
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   0
            Width           =   3768
            Begin XtremeSuiteControls.RadioButton ContructType 
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   35
               Top             =   240
               Width           =   735
               _Version        =   786432
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " —ﬂÌ»"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton ContructType 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   735
               _Version        =   786432
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "’Ì«‰…"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·⁄ﬁœ"
               Height          =   285
               Index           =   14
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   240
               Width           =   810
            End
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   288
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   3096
            _ExtentX        =   5450
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DBContractDate 
            Height          =   315
            Left            =   5280
            TabIndex        =   51
            Top             =   705
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102760449
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DBEndGuranteeDate 
            Height          =   315
            Left            =   5550
            TabIndex        =   52
            Top             =   1200
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102760449
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DueDate 
            Height          =   315
            Left            =   10140
            TabIndex        =   53
            Top             =   1200
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102760449
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbEmployee 
            Height          =   288
            Left            =   120
            TabIndex        =   54
            Top             =   1200
            Width           =   4536
            _ExtentX        =   7990
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   285
            Index           =   29
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   780
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„— Ã⁄"
            Height          =   285
            Index           =   28
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   720
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ/«·›« Ê—…"
            Height          =   285
            Index           =   0
            Left            =   8700
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   192
            Index           =   3
            Left            =   12768
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   276
            Width           =   516
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   288
            Index           =   1
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   720
            Width           =   456
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ "
            Height          =   285
            Index           =   4
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   405
            Index           =   5
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   780
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  ≈‰ Â«¡ «·÷„«‰"
            Height          =   288
            Index           =   6
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1200
            Width           =   1176
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   288
            Index           =   7
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1680
            Width           =   696
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·’Ì«‰…"
            Height          =   285
            Index           =   8
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1680
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’Ì«‰…"
            Height          =   285
            Index           =   9
            Left            =   12060
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   285
            Index           =   11
            Left            =   12060
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰œÊ»"
            Height          =   288
            Index           =   12
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1200
            Width           =   576
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   2190
         Left            =   14130
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   45
         Width           =   13395
         _cx             =   23627
         _cy             =   3863
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
         Begin VB.TextBox StructureType 
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
            Height          =   324
            Left            =   8520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   960
            Width           =   3600
         End
         Begin VB.TextBox ContractValidity 
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
            Left            =   8520
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   1320
            Width           =   3600
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Height          =   600
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1320
            Width           =   3048
            Begin XtremeSuiteControls.RadioButton SecondEntrance 
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   92
               Top             =   240
               Width           =   615
               _Version        =   786432
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "»œÊ‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton SecondEntrance 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Width           =   735
               _Version        =   786432
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊÃÊœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œŒ· À«‰Ì"
               Height          =   285
               Index           =   25
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   735
            Width           =   3048
            Begin XtremeSuiteControls.RadioButton additionalRoom 
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   88
               Top             =   240
               Width           =   615
               _Version        =   786432
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "»œÊ‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton additionalRoom 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   89
               Top             =   240
               Width           =   735
               _Version        =   786432
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊÃÊœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "€—›…"
               Height          =   285
               Index           =   24
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   120
            Width           =   3048
            Begin XtremeSuiteControls.RadioButton RescueDevice 
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   84
               Top             =   240
               Width           =   615
               _Version        =   786432
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "»œÊ‰"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RescueDevice 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   85
               Top             =   240
               Width           =   735
               _Version        =   786432
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊÃÊœ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÃÂ«“ «‰ﬁ«–"
               Height          =   285
               Index           =   23
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.TextBox VVVF 
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
            Height          =   324
            Left            =   3612
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   960
            Width           =   3570
         End
         Begin VB.TextBox Capacity 
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
            Left            =   3612
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   600
            Width           =   3570
         End
         Begin VB.TextBox DoorOpening 
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
            Height          =   300
            Left            =   3612
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   252
            Width           =   3570
         End
         Begin VB.TextBox Design 
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
            Left            =   8535
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1680
            Width           =   3588
         End
         Begin VB.TextBox Color 
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
            Left            =   3600
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1320
            Width           =   3555
         End
         Begin VB.TextBox DoorSize 
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
            Height          =   324
            Left            =   8535
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   960
            Width           =   3588
         End
         Begin VB.TextBox NoOfStops 
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
            Left            =   8535
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   600
            Width           =   3588
         End
         Begin VB.TextBox NoOfUnits 
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
            Height          =   300
            Left            =   8535
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   252
            Width           =   3588
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„‰‘√…"
            Height          =   300
            Index           =   27
            Left            =   12255
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   960
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·⁄ﬁœ"
            Height          =   285
            Index           =   26
            Left            =   12375
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "VVVF"
            Height          =   300
            Index           =   22
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄…"
            Height          =   285
            Index           =   21
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÂ… «·»«» "
            Height          =   270
            Index           =   20
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ’„Ì„"
            Height          =   285
            Index           =   19
            Left            =   12375
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«··Ê‰"
            Height          =   285
            Index           =   18
            Left            =   7455
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÕÃ„ «·»«»"
            Height          =   300
            Index           =   17
            Left            =   12255
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   960
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Êﬁ›« "
            Height          =   285
            Index           =   16
            Left            =   12255
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊÕœ« "
            Height          =   270
            Index           =   15
            Left            =   12255
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   255
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "FrmOldContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Public ScrenFlg As Integer
Dim mRow As Integer

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

Label1(10).Caption = "Type"
RdType(0).RightToLeft = False
RdType(1).RightToLeft = False
RdType(0).Caption = "Cont."
RdType(1).Caption = "Bills"
Label1(11).Caption = "Due Date"
Label1(12).Caption = "Employee"

    Me.Caption = "Old Contract Data"
    Label1(2).Caption = Me.Caption

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        If ScrenFlg = 1 Then
        .TextMatrix(0, .ColIndex("CusID")) = "Supplier Name"
        Else
        .TextMatrix(0, .ColIndex("CusID")) = "Customer Name"
        End If
        .TextMatrix(0, .ColIndex("ContractNo")) = "Contract No."
        .TextMatrix(0, .ColIndex("ContractDate")) = "Contract Date"
.TextMatrix(0, .ColIndex("ContractValue")) = "Contract Value"
.TextMatrix(0, .ColIndex("EndGuranteeDate")) = "End Gurantee Date"
.TextMatrix(0, .ColIndex("NetValue")) = "Maintenance Value"
.TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With
    Label1(0).Caption = "Cont.No"
    Label1(4).Caption = "Contr.Date"
    If ScrenFlg = 1 Then
    Label1(1).Caption = "Supplier"
    Else
    Label1(1).Caption = "Customer"
    End If
    Label1(6).Caption = "End Gurantee"
    Label1(9).Caption = "Type"
     Label1(5).Caption = "Maint. Value"
    Label1(8).Caption = "Contract Value"
    Label1(7).Caption = "Remarks"
XPPnlTime.Caption = "Print End Contruct"
    Label1(3).Caption = "ID"
    lbl(1).Caption = "From"
 lbl(0).Caption = "To"
    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."
BtnPrint.Caption = "Print"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

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

'    If DoPremis(Do_Delete, Me.name, True) = False Then
'        Exit Sub
'    End If

    If TxtVac_ID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
        MSGType = MsgBox("ÂConfirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
     End If

        If MSGType = vbYes Then
        Cn.Execute " delete  from Transactions where OldContID=" & val(TxtVac_ID.Text) & ""
            RsSavRec.find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry... Can not Delete.  is related to with other data"
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
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

'    If DoPremis(Do_Edit, Me.name, True) = False Then
'        Exit Sub
'    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtContractNo.SetFocus
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
            Msg = Msg & " ·Currently can not be edited" & CHR(13)
            Msg = Msg & "Where it was being edited by another user on the network"
           
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

'    If DoPremis(Do_New, Me.name, True) = False Then
'        Exit Sub
'    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    '-----------------------------------
    Me.TxtVac_ID.Text = ""
    Me.TxtContractNo.Text = ""
    Me.DBCboClientName.BoundText = ""
    txtContractValue.Text = ""
    TxtRemarks.Text = ""
    DueDate.value = Date
    '-----------------------------------
    TxtModFlg.Text = "N"

    My_SQL = "TblOLDContract"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If
    TxtVac_ID.Text = TxtSerial.Text
    'ContructType_Click
    
    rs.Close
    CmbType.ListIndex = 0
    DBCboClientName.SetFocus
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
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrint_Click()
        Dim xApp As New CRAXDRT.Application

    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report

    Dim rs As ADODB.Recordset
    Dim cCompanyInfo As ClsCompanyInfo
    Set cCompanyInfo = New ClsCompanyInfo
 '   sql = "SELECT * from projects where 1=1"
 Dim sql As String
 
 sql = "SELECT     TOP 100 PERCENT dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblOLDContract.ContractNo, "
sql = sql & "                      dbo.TblOLDContract.ContractValue, dbo.TblOLDContract.ContractDate, dbo.TblOLDContract.EndGuranteeDate, dbo.TblOLDContract.Remarks,"
sql = sql & "                      dbo.TblOLDContract.netvalue , dbo.TblOLDContract.PaymentType, dbo.TblOLDContract.Vlue,dbo.TblOLDContract.ReturnNo,dbo.TblOLDContract.ReturnValue"
sql = sql & " FROM         dbo.TblOLDContract INNER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblOLDContract.CusID = dbo.TblCustemers.CusID"
sql = sql & "   where 1=1"
    
    If Not IsNull(XPDtbFrom.value) Then
    sql = sql & "  and EndGuranteeDate >=  " & SQLDate(XPDtbFrom, True)
 
    End If
    
        If Not IsNull(XPDtpTo.value) Then
    sql = sql & "  and EndGuranteeDate<=  " & SQLDate(XPDtpTo, True)
 
    End If
 
sql = sql & "   ORDER BY dbo.TblOLDContract.EndGuranteeDate "

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\REPORTS NEW\OldContract.rpt")
    Else
    
        Set xReport = xApp.OpenReport(App.path & "\reports\REPORTS NEW\OldContract.rpt")
    End If

   
    xReport.Database.SetDataSource rs
     Dim cAccountReport As New ClsReportViewer
    Dim FrmReport As New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.TxtPath = (App.path & "\reports\REPORTS NEW\OldContract.rpt")
    xReport.reporttitle = "  «·⁄ﬁÊœ «·”«»ﬁ…"
    
       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
       xReport.ParameterFields(2).AddCurrentValue user_name
       
       xReport.ParameterFields(3).AddCurrentValue Format(XPDtbFrom.value, "DD/MM/YYYY")
        xReport.ParameterFields(4).AddCurrentValue Format(XPDtpTo, "DD/MM/YYYY")
      
      
    FrmReport.CRViewer.viewReport
 cAccountReport.CreateLogo xReport
       
      
   'xReport.reporttitle = cCompanyInfo.ArabCompanyName
          
    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
'    SendKeys "{RIGHT}"
          
End Sub

Private Sub btnQuery_Click()

 FrmCommisSearch.mType = 1
 FrmCommisSearch.Caption = "»ÕÀ" & Label1(2).Caption
 FrmCommisSearch.show
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
'                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
'                CtrlTxt.SetFocus
           '     Exit Sub
            End If
        End If

    Next
    
    If TxtContractNo.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« Ì—ÃÏ «œŒ«· —ﬁ„ «·⁄ﬁœ", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please, Enter the contract No.", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    If txtContractValue.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« Ì—ÃÏ «œŒ«· ﬁÌ„… «·⁄ﬁœ", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please, Enter the contract Value", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblOLDContract", "ContractNo", Trim(TxtContractNo.Text), "ContractNo", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
        Msg = "This type already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtContractNo.SetFocus
    
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
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
   Else
   MsgBox "Sorry...error in douring enter data", vbOKOnly + vbMsgBoxRight, App.title
   End If

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
If SystemOptions.UserInterface = ArabicInterface Then
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
Else
    If FristCount = LastCount Then
        Msg = "No new data found"
    Else
        Msg = "No Rec.Before Update " & vbCrLf & FristCount & vbCrLf & "No Rec.After Update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "No New Record" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "No Deleted Record"" & vbCrLf & FristCount - LastCount"
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub ChkType()
    If ContructType(0).value = True Then
        Label1(15).Caption = " «·ÊÕœ« "
        Label1(16).Caption = "⁄œœ «·Êﬁ›« "
        Label1(17).Visible = True
        Label1(18).Visible = True
        Label1(19).Visible = True
        Label1(20).Visible = True
        Label1(21).Visible = True
        Label1(22).Visible = True
        DoorSize.Visible = True
        ContractValidity.Visible = True
        Design.Visible = True
        DoorOpening.Visible = True
        Capacity.Visible = True
        VVVF.Visible = True
        Color.Visible = True
        Label1(27).Visible = False
        StructureType.Visible = False
        Frame3.Visible = True
        Frame4.Visible = True
        Frame5.Visible = True
    ElseIf ContructType(1).value = True Then
        Label1(15).Caption = " «—ÌŒ «·»œ«Ì…"
        Label1(16).Caption = " «—ÌŒ «·‰Â«Ì…"
        Label1(17).Visible = False
        Label1(18).Visible = False
        Label1(19).Visible = False
        Label1(20).Visible = False
        Label1(21).Visible = False
        Label1(22).Visible = False
        DoorSize.Visible = False
        ContractValidity.Visible = False
        Design.Visible = False
        DoorOpening.Visible = False
        Capacity.Visible = False
        VVVF.Visible = False
        Color.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
        Label1(27).Visible = True
        StructureType.Visible = True
    End If
End Sub



Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCountriesNames Me.DBCboClientName
    End If

End Sub

Private Sub DcbPaymentType_Change()
If val(DcbPaymentType.ListIndex) = 0 Then
TxtNetValue.Text = TxtValue.Text
Label1(8).Caption = "ﬁÌ„… «·’Ì«‰…"
Else
TxtNetValue.Text = (val(TxtValue.Text) * val(txtContractValue.Text)) / 100
TxtNetValue.Text = Round(val(TxtNetValue.Text), 2)
Label1(8).Caption = "‰”»… «·’Ì«‰…"
End If
End Sub

Private Sub DcbPaymentType_Click()
DcbPaymentType_Change
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
If SystemOptions.UserInterface = ArabicInterface Then
With DcbPaymentType
.Clear
.AddItem "ﬁÌ„…"
.AddItem "‰”»…"
End With
Else
With DcbPaymentType
.Clear
.AddItem "Vlue"
.AddItem "Percentage"
End With
End If
    Me.XPDtbFrom = FirstPeriodDateInthisYear
    Me.XPDtpTo = Date


    My_SQL = "select * from TblOLDContract "
    If ScrenFlg = 1 Then
    My_SQL = My_SQL & " where ScrenFlg=1"
    Else
    My_SQL = My_SQL & " where ScrenFlg is null"
    End If
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    'RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    If ScrenFlg = 1 Then
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
    Else
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    End If
    Dcombos.GetSalesRepData Me.DcbEmployee
    Set cSearch = New clsDCboSearch
    Set cSearch.Client = Me.DBCboClientName

    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("CusID"), Me.DBCboClientName

    FillGridWithData

    With Me.Grid
'        .Cell(flexcpPicture, 0, .ColIndex("ContractNo")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
'        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

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
    
    ChkType

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
   ' Set FrmVacancy = Nothing

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

    Set cSearch = Nothing
ErrTrap:
End Sub

Private Sub Form_Activate()
If SystemOptions.UserInterface = ArabicInterface Then
 If ScrenFlg = 1 Then
 Label1(1).Caption = "«·„Ê—œ"
 Grid.TextMatrix(0, Grid.ColIndex("CusID")) = "«·„Ê—œ"
 Label1(2).Caption = "»Ì«‰«  «·⁄ﬁÊœ/Ê›Ê« Ì— «·„‘ —Ì«  «·ﬁœÌ„…"
 Else
 Label1(2).Caption = "»Ì«‰«  «·⁄ﬁÊœ/Ê›Ê« Ì— «·„»Ì⁄«  «·ﬁœÌ„…"
 Grid.TextMatrix(0, Grid.ColIndex("CusID")) = "«·⁄„Ì·"
 Label1(1).Caption = "«·⁄„Ì·"
 End If
 End If
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblOLDContract", "id", "")
    RsSavRec.AddNew
    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Sub SaveBill(Optional OldContID As Double, Optional OldValue As Double)
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
Dim StrRecID As String
sql = "Select * from Transactions where 1=-1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
StrRecID = new_id("Transactions", "Transaction_ID", "")
Rs3.AddNew
Rs3("Transaction_ID").value = StrRecID
If ScrenFlg = 1 Then
Rs3("Transaction_Type").value = 73
Else
Rs3("Transaction_Type").value = 71
End If
Rs3("Transaction_Date").value = DBContractDate.value
Rs3("DueDate").value = DueDate.value
Rs3("CusID").value = val(DBCboClientName.BoundText)
Rs3("BranchId").value = Current_branch
Rs3("NoteSerial1").value = TxtContractNo.Text
Rs3("ManualNO").value = TxtContractNo.Text
Rs3("OldContID").value = OldContID
Rs3("OldValue").value = OldValue
Rs3("Emp_ID").value = val(DcbEmployee.BoundText)
Rs3("PaymentType").value = 1
Rs3.update
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap
    
    If Me.TxtModFlg.Text = "E" Then
    Cn.Execute " delete  from Transactions where OldContID=" & val(TxtVac_ID.Text) & ""
    End If
    
     RsSavRec.Fields("DueDate").value = DueDate.value
     RsSavRec.Fields("CusID").value = IIf(val(DBCboClientName.BoundText) <> 0, val(DBCboClientName.BoundText), Null)
     RsSavRec.Fields("ContractNo").value = IIf(TxtContractNo.Text <> "", Trim(TxtContractNo.Text), Null)
     RsSavRec.Fields("ContractDate").value = DBContractDate.value
     RsSavRec.Fields("ContractValue").value = IIf(txtContractValue.Text <> "", Trim(txtContractValue.Text), Null)
     RsSavRec.Fields("EndGuranteeDate").value = DBEndGuranteeDate.value
     RsSavRec.Fields("Remarks").value = IIf(TxtRemarks.Text <> "", Trim(TxtRemarks.Text), Null)
     RsSavRec.Fields("PaymentType").value = IIf(DcbPaymentType.ListIndex <> -1, val(DcbPaymentType.ListIndex), Null)
     RsSavRec.Fields("Vlue").value = IIf(TxtValue.Text <> "", val(TxtValue.Text), Null)
     RsSavRec.Fields("NetValue").value = IIf(TxtNetValue.Text <> "", val(TxtNetValue.Text), Null)
     RsSavRec.Fields("Emp_ID").value = IIf(val(Me.DcbEmployee.BoundText) <> 0, val(DcbEmployee.BoundText), Null)
     
     If RdType(1).value = True Then
     RsSavRec.Fields("RdType").value = 1
     ElseIf RdType(0).value = True Then
     RsSavRec.Fields("RdType").value = 0
     ElseIf RdType(2).value = True Then
     RsSavRec.Fields("RdType").value = 2
     ElseIf RdType(3).value = True Then
     RsSavRec.Fields("RdType").value = 3
     End If
     
    '#######################################################################################
    
    RsSavRec.Fields("NoOfUnits").value = IIf(NoOfUnits.Text <> "", NoOfUnits.Text, Null)
    RsSavRec.Fields("NoOfStops").value = IIf(NoOfStops.Text <> "", NoOfStops.Text, Null)
    RsSavRec.Fields("DoorSize").value = IIf(DoorSize.Text <> "", DoorSize.Text, Null)
    RsSavRec.Fields("Color").value = IIf(Color.Text <> "", Color.Text, Null)
    RsSavRec.Fields("Design").value = IIf(Design.Text <> "", Design.Text, Null)
    RsSavRec.Fields("DoorOpening").value = IIf(DoorOpening.Text <> "", DoorOpening.Text, Null)
    RsSavRec.Fields("Capacity").value = IIf(Capacity.Text <> "", Capacity.Text, Null)
    RsSavRec.Fields("VVVF").value = IIf(VVVF.Text <> "", VVVF.Text, Null)
    RsSavRec.Fields("ContractValidity").value = IIf(ContractValidity.Text <> "", ContractValidity.Text, Null)
    RsSavRec.Fields("StructureType").value = IIf(StructureType.Text <> "", StructureType.Text, Null)
    RsSavRec.Fields("ReturnNo").value = IIf(Txt_ReturnNo.Text <> "", Trim(Txt_ReturnNo.Text), Null)
    RsSavRec.Fields("ReturnValue").value = IIf(Txt_ReturnValue.Text <> "", Trim(Txt_ReturnValue.Text), Null)
    
    
    If ContructType(1).value = True Then
        RsSavRec.Fields("ContructType").value = 1
    Else
        RsSavRec.Fields("ContructType").value = 0
    End If
    
    If RescueDevice(1).value = True Then
        RsSavRec.Fields("RescueDevice").value = 1
    Else
        RsSavRec.Fields("RescueDevice").value = 0
    End If
    
    If additionalRoom(1).value = True Then
        RsSavRec.Fields("additionalRoom").value = 1
    Else
        RsSavRec.Fields("additionalRoom").value = 0
    End If
    
    If SecondEntrance(1).value = True Then
        RsSavRec.Fields("SecondEntrance").value = 1
    Else
        RsSavRec.Fields("SecondEntrance").value = 0
    End If

    
    '######################################################################################
    
     If ScrenFlg = 1 Then
     RsSavRec.Fields("ScrenFlg").value = 1
    Else
    RsSavRec.Fields("ScrenFlg").value = Null
    End If
    RsSavRec.update
    If RdType(1).value = True Then
    SaveBill val(TxtVac_ID.Text), val(txtContractValue.Text)
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
    MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGridWithData
    TxtModFlg = "R"
    If mRow > 0 Then
        FindRec val(Me.Grid.TextMatrix(mRow, Me.Grid.ColIndex("id")))
     
        FiLLTXT
    End If
    ChkType
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
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtSerial.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    DueDate.value = IIf(IsNull(RsSavRec.Fields("DueDate").value), Date, RsSavRec.Fields("DueDate").value)
    TxtContractNo.Text = IIf(IsNull(RsSavRec.Fields("ContractNo").value), "", RsSavRec.Fields("ContractNo").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("REMARKS").value), "", RsSavRec.Fields("REMARKS").value)
    DBContractDate.value = IIf(IsNull(RsSavRec.Fields("ContractDate").value), Date, RsSavRec.Fields("ContractDate").value)
    DBEndGuranteeDate.value = IIf(IsNull(RsSavRec.Fields("EndGuranteeDate").value), Date, RsSavRec.Fields("EndGuranteeDate").value)
    txtContractValue.Text = IIf(IsNull(RsSavRec.Fields("ContractValue").value), "", RsSavRec.Fields("ContractValue").value)
    TxtNetValue.Text = IIf(IsNull(RsSavRec.Fields("NetValue").value), 0, RsSavRec.Fields("NetValue").value)
    TxtValue.Text = IIf(IsNull(RsSavRec.Fields("Vlue").value), 0, RsSavRec.Fields("Vlue").value)
    DcbPaymentType.ListIndex = IIf(IsNull(RsSavRec.Fields("PaymentType").value), -1, RsSavRec.Fields("PaymentType").value)
    Me.DcbEmployee.BoundText = IIf(IsNull(RsSavRec.Fields("Emp_ID").value), "", RsSavRec.Fields("Emp_ID").value)
    
    If Not IsNull(RsSavRec.Fields("RdType").value) Then
     If RsSavRec.Fields("RdType").value = 1 Then
     RdType(1).value = True
     ElseIf RsSavRec.Fields("RdType").value = 0 Then
     RdType(0).value = True
     ElseIf RsSavRec.Fields("RdType").value = 2 Then
     RdType(2).value = True
     ElseIf RsSavRec.Fields("RdType").value = 3 Then
     RdType(3).value = True
     End If
    End If
    
        
    
    '#############################################################################################################################
    NoOfUnits.Text = IIf(IsNull(RsSavRec.Fields("NoOfUnits").value), "", RsSavRec.Fields("NoOfUnits").value)
    NoOfStops.Text = IIf(IsNull(RsSavRec.Fields("NoOfStops").value), "", RsSavRec.Fields("NoOfStops").value)
    DoorSize.Text = IIf(IsNull(RsSavRec.Fields("DoorSize").value), "", RsSavRec.Fields("DoorSize").value)
    Color.Text = IIf(IsNull(RsSavRec.Fields("Color").value), "", RsSavRec.Fields("Color").value)
    Design.Text = IIf(IsNull(RsSavRec.Fields("Design").value), "", RsSavRec.Fields("Design").value)
    DoorOpening.Text = IIf(IsNull(RsSavRec.Fields("DoorOpening").value), "", RsSavRec.Fields("DoorOpening").value)
    Capacity.Text = IIf(IsNull(RsSavRec.Fields("Capacity").value), "", RsSavRec.Fields("Capacity").value)
    VVVF.Text = IIf(IsNull(RsSavRec.Fields("VVVF").value), "", RsSavRec.Fields("VVVF").value)
    ContractValidity.Text = IIf(IsNull(RsSavRec.Fields("ContractValidity").value), "", RsSavRec.Fields("ContractValidity").value)
    StructureType.Text = IIf(IsNull(RsSavRec.Fields("StructureType").value), "", RsSavRec.Fields("StructureType").value)
    Txt_ReturnNo.Text = IIf(IsNull(RsSavRec.Fields("ReturnNo").value), "", RsSavRec.Fields("ReturnNo").value)
    Txt_ReturnValue.Text = IIf(IsNull(RsSavRec.Fields("ReturnValue").value), "", RsSavRec.Fields("ReturnValue").value)
    
    
    
    If Not IsNull(RsSavRec.Fields("ContructType").value) Then
        If (RsSavRec.Fields("ContructType").value) = 1 Then
            ContructType(1).value = True
         Else
            ContructType(0).value = True
        End If
    Else
        ContructType(0).value = True
    End If
    
    If Not IsNull(RsSavRec.Fields("RescueDevice").value) Then
        If (RsSavRec.Fields("RescueDevice").value) = 1 Then
            RescueDevice(1).value = True
         Else
            RescueDevice(0).value = True
        End If
    Else
        RescueDevice(0).value = True
    End If
    
    If Not IsNull(RsSavRec.Fields("additionalRoom").value) Then
        If (RsSavRec.Fields("additionalRoom").value) = 1 Then
            additionalRoom(1).value = True
         Else
            additionalRoom(0).value = True
        End If
    Else
        additionalRoom(0).value = True
    End If
    
    If Not IsNull(RsSavRec.Fields("SecondEntrance").value) Then
        If (RsSavRec.Fields("SecondEntrance").value) = 1 Then
            SecondEntrance(1).value = True
         Else
            SecondEntrance(0).value = True
        End If
    Else
        SecondEntrance(0).value = True
    End If
    
    '#############################################################################################################################
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With
    
    ChkType

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_Click()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
    mRow = Me.Grid.Row
    FiLLTXT
    ChkType
ErrTrap:
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub



Private Sub RdType_Click(Index As Integer)
If RdType(2).value = True Then
Label1(28).Visible = True
Label1(28).Caption = "—ﬁ„ «·„— Ã⁄"
Label1(29).Visible = True
Label1(29).Caption = "«·ﬁÌ„…"
Txt_ReturnNo.Visible = True
Txt_ReturnValue.Visible = True
ElseIf RdType(3).value = True Then
Label1(28).Visible = True
Label1(28).Caption = "—ﬁ„ «·”‰œ"
Label1(29).Visible = True
Label1(29).Caption = "«·ﬁÌ„…"
Txt_ReturnNo.Visible = True
Txt_ReturnValue.Visible = True
Else
Label1(28).Visible = False
Label1(29).Visible = False
Txt_ReturnNo.Visible = False
Txt_ReturnValue.Visible = False
End If


End Sub

Private Sub TXtContractValue_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.TXtContractValue.Text, 1)
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "id=" & RecId, , adSearchForward, 1

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
    My_SQL = "select * From TblOLDContract"
    If ScrenFlg = 1 Then
    My_SQL = My_SQL & " where   ScrenFlg=1"
    Else
    My_SQL = My_SQL & " where   ScrenFlg Is null"
    End If
    My_SQL = My_SQL & " order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
               
               .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
           If Not (IsNull(rs.Fields("PaymentType").value)) Then
           If rs.Fields("PaymentType").value = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("PaymentType")) = "ﬁÌ„…"
           Else
           .TextMatrix(i, .ColIndex("PaymentType")) = "Value"
           End If
           ElseIf rs.Fields("PaymentType").value = 1 Then
              If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("PaymentType")) = "‰”»…"
           Else
           .TextMatrix(i, .ColIndex("PaymentType")) = "Percentage"
           End If
           End If
           End If
           .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
           .TextMatrix(i, .ColIndex("Vlue")) = IIf(IsNull(rs.Fields("Vlue").value), 0, rs.Fields("Vlue").value)
           .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(rs.Fields("NetValue").value), 0, rs.Fields("NetValue").value)
           
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value)
                .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
                .TextMatrix(i, .ColIndex("ContractDate")) = IIf(IsNull(rs.Fields("ContractDate").value), "", rs.Fields("ContractDate").value)
                .TextMatrix(i, .ColIndex("ContractValue")) = IIf(IsNull(rs.Fields("ContractValue").value), "", rs.Fields("ContractValue").value)
                .TextMatrix(i, .ColIndex("EndGuranteeDate")) = IIf(IsNull(rs.Fields("EndGuranteeDate").value), "", rs.Fields("EndGuranteeDate").value)
               .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs.Fields("Remarks").value), "", rs.Fields("Remarks").value)
               .TextMatrix(i, .ColIndex("ReturnNo")) = IIf(IsNull(rs.Fields("ReturnNo").value), "", rs.Fields("ReturnNo").value)
               .TextMatrix(i, .ColIndex("ReturnValue")) = IIf(IsNull(rs.Fields("ReturnValue").value), "", rs.Fields("ReturnValue").value)
               
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

Private Function CheckDelCountry(LngCusID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & LngCusID & ""
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

Private Sub TxtValue_Change()
DcbPaymentType_Change
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue.Text, 0)
End Sub
