VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmKhsm 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·Œ’Ê„« "
   ClientHeight    =   7215
   ClientLeft      =   2685
   ClientTop       =   2475
   ClientWidth     =   11955
   HelpContextID   =   540
   Icon            =   "FrmKhsm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   11955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1005
      Left            =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6135
      Width           =   11850
      _cx             =   20902
      _cy             =   1773
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
         Height          =   480
         Left            =   8235
         TabIndex        =   7
         Top             =   495
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":038A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   480
         Left            =   6255
         TabIndex        =   6
         Top             =   495
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":0724
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   480
         Left            =   7215
         TabIndex        =   8
         Top             =   495
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":0ABE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   480
         Left            =   5205
         TabIndex        =   9
         Top             =   495
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":0E58
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   480
         Left            =   4365
         TabIndex        =   10
         Top             =   495
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":11F2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5010
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   1035
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
         ButtonImage     =   "FrmKhsm.frx":178C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   3960
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   1035
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
         ButtonImage     =   "FrmKhsm.frx":1B26
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   2940
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1065
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
         ButtonImage     =   "FrmKhsm.frx":1EC0
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   480
         Left            =   45
         TabIndex        =   11
         Top             =   495
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":225A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin MSDataListLib.DataCombo DCUser 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   9000
         TabIndex        =   54
         Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
         Top             =   240
         Width           =   1830
         _ExtentX        =   3228
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
      Begin ImpulseButton.ISButton ISButton1 
         Height          =   480
         Left            =   960
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   495
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   847
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…„–ﬂ—… Œ’„ "
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
         ButtonImage     =   "FrmKhsm.frx":25F4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   480
         Left            =   2640
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   495
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":8E56
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton8 
         Height          =   480
         Left            =   3360
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   495
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   847
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
         ButtonImage     =   "FrmKhsm.frx":F6B8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ—— »Ê«”ÿ…"
         Height          =   300
         Index           =   13
         Left            =   10860
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   270
         Width           =   915
      End
      Begin VB.Label LBLWhereSTR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   255
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   -600
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label LBLhOURrATE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   255
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   -960
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   135
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   135
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   2
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   135
         Width           =   975
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   -75
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   -90
      Width           =   12060
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3795
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Text            =   "modflag"
         Top             =   165
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtKhsmEdafa_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   2805
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   165
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   4080
         Top             =   60
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
               Picture         =   "FrmKhsm.frx":FA52
               Key             =   "Emp_Name"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmKhsm.frx":FDEC
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmKhsm.frx":10186
               Key             =   "Emp_Code"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmKhsm.frx":10720
               Key             =   "Emp_Salary"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   210
         TabIndex        =   17
         Top             =   195
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
         ButtonImage     =   "FrmKhsm.frx":10ABA
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   675
         TabIndex        =   16
         Top             =   195
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
         ButtonImage     =   "FrmKhsm.frx":10E54
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   15
         Top             =   195
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
         ButtonImage     =   "FrmKhsm.frx":111EE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Top             =   195
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
         ButtonImage     =   "FrmKhsm.frx":11588
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2040
         Picture         =   "FrmKhsm.frx":11922
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ’Ê„« "
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
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   90
         Width           =   1680
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5640
      Left            =   -60
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   600
      Width           =   12015
      Begin XtremeSuiteControls.CheckBox ChAccept 
         Height          =   255
         Left            =   -240
         TabIndex        =   61
         Top             =   1920
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "„Ê›ﬁ… «·„ÊŸ›"
         ForeColor       =   16711680
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.ComboBox CmbYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1380
         Width           =   1425
      End
      Begin VB.TextBox TxtVal 
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
         Left            =   10110
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox TxtResonalarm 
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
         Height          =   2940
         Left            =   6240
         MaxLength       =   50
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox TxtAlrmOrder 
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
         Left            =   3030
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   75
         Width           =   1545
      End
      Begin VB.TextBox TxtSearchCode 
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
         TabIndex        =   39
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox CboCalType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8055
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1380
         Width           =   2880
      End
      Begin VB.ComboBox CmbMonth 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2445
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1380
         Width           =   1665
      End
      Begin MSDataListLib.DataCombo DCEmp_Name 
         Height          =   315
         Left            =   6255
         TabIndex        =   2
         Top             =   540
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
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
         Left            =   9000
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· —ﬁ„ «·Œ’„"
         Top             =   75
         Width           =   1935
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
         Left            =   6240
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Text            =   "1"
         Top             =   1380
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   330
         Left            =   6255
         TabIndex        =   1
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   12648447
         Format          =   94044161
         CurrentDate     =   38887
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
         Height          =   2940
         Left            =   360
         MaxLength       =   50
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2280
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   540
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSanction 
         Height          =   315
         Left            =   6240
         TabIndex        =   49
         Top             =   960
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCComponent 
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DateSet 
         Height          =   315
         Left            =   2445
         TabIndex        =   63
         Top             =   1800
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         Format          =   94044161
         CurrentDate     =   38887
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÊÌ…"
         Height          =   195
         Index           =   3
         Left            =   4155
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1860
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
         Height          =   615
         Index           =   4
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã“«¡"
         Height          =   195
         Index           =   14
         Left            =   11355
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”»» «·«‰–«—"
         Height          =   615
         Index           =   12
         Left            =   11040
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€"
         Height          =   255
         Index           =   11
         Left            =   10920
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1830
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   195
         Index           =   10
         Left            =   5775
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1830
         Width           =   4170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„"
         Height          =   255
         Index           =   9
         Left            =   4710
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï «‰–«— —ﬁ„"
         Height          =   255
         Index           =   8
         Left            =   4740
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ «·Œ’„"
         Height          =   255
         Index           =   7
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Œ’„"
         Height          =   255
         Index           =   5
         Left            =   10890
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1410
         Width           =   870
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "”‰…"
         Height          =   195
         Index           =   2
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1380
         Width           =   270
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   195
         Index           =   1
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1380
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì‰›– ›Ï "
         Height          =   195
         Index           =   0
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1395
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌÊ„"
         Height          =   195
         Index           =   3
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1410
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Œ’„"
         Height          =   195
         Index           =   0
         Left            =   7950
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   135
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Œ’„"
         Height          =   195
         Index           =   6
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   195
         Index           =   0
         Left            =   11265
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œ… «·Œ’„"
         Height          =   255
         Index           =   1
         Left            =   6930
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1410
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmKhsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim RecId As String
Dim II As Long
Dim StrDate As Date
Dim cSearch As clsDCboSearch

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
Dim StrSQL As String
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtKhsmEdafa_ID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.title)
       Else
       MSGType = MsgBox("Confirm Delete", vbYesNo + vbMsgBoxRight, App.title)
     End If
        If MSGType = vbYes Then
            RsSavRec.find "KhsmEdafa_ID=" & val(TxtKhsmEdafa_ID.Text), , adSearchForward, 1
            RsSavRec.delete
            StrSQL = "Delete From TblInforVacatiom Where KhaseID=" & val(TxtKhsmEdafa_Code.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where KsmID=" & val(Me.TxtKhsmEdafa_Code.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegister Where KsmID=" & val(Me.TxtKhsmEdafa_Code.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
         If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Delete Successfully", vbOKOnly + vbMsgBoxRight, App.title
            End If
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
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"
    Me.DCUser.BoundText = user_id
    My_SQL = "select * From tblKhsmEdafa where KhsmEdafa_Type=0"
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

    Dim Msg As String
    'Dim StrVacName As String
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

    If DCEmp_Name.BoundText = "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "⁄›Ê« Ì—ÃÏ  ÕœÌœ «”„ «·„ÊŸ›"
       Else
       Msg = "Please slect Employee"
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

    Dim FristCount As Long
    Dim Msg As String
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

Private Sub DcbSanction_Change()
DcbSanction_Click (0)
End Sub

Private Sub DcbSanction_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(DcbSanction.BoundText) <> 0 Then
GetMofrSan val(DcbSanction.BoundText)
End If
End If
End Sub
Sub GetEmployeeWarning(Optional ID As Integer = 0)
Dim Rs6 As ADODB.Recordset
Dim sql As String
Set Rs6 = New ADODB.Recordset
sql = "SELECT    * "
sql = sql & " From dbo.TblEmployeeWarrning"
sql = sql & " Where (ID = " & ID & ")"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
DCEmp_Name.BoundText = IIf(IsNull(Rs6("Emp_ID").value), 0, Rs6("Emp_ID").value)
DcbSanction.BoundText = IIf(IsNull(Rs6("SanctionID").value), 0, Rs6("SanctionID").value)
DcbDept.BoundText = IIf(IsNull(Rs6("DeptID").value), 0, Rs6("DeptID").value)
TxtResonalarm.Text = IIf(IsNull(Rs6("Remark").value), "", Rs6("Remark").value)
Else
TxtResonalarm.Text = ""
DcbSanction.BoundText = 0
DcbDept.BoundText = 0
DCComponent.BoundText = 0
End If
End Sub

Sub GetMofrSan(Optional ID As Integer = 0)
Dim Rs6 As ADODB.Recordset
Dim sql As String
Set Rs6 = New ADODB.Recordset
sql = "SELECT    Mofrd "
sql = sql & " From dbo.TblAdminSanction"
sql = sql & " Where (ID = " & ID & ")"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
DCComponent.BoundText = IIf(IsNull(Rs6("Mofrd").value), 0, Rs6("Mofrd").value)
Else
DCComponent.BoundText = 0
End If
End Sub
Private Sub DCComponent_Change()
DCComponent_Click (0)
End Sub

Private Sub DCComponent_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
CboCalType.ListIndex = -1
RetrivMofrd val(DCComponent.BoundText)
End If
End Sub
Sub RetrivMofrd(Optional MofrID As Integer)
Dim Rs7 As ADODB.Recordset
Dim sql As String
If MofrID <> 0 Then
Set Rs7 = New ADODB.Recordset
sql = " SELECT     id, Unit"
sql = sql & " From dbo.MOFRAD"
sql = sql & " Where (id = " & MofrID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
If Not (IsNull(Rs7("Unit").value)) Then
If Rs7("Unit").value = 0 Then
CboCalType.ListIndex = 1
ElseIf Rs7("Unit").value = 1 Then
CboCalType.ListIndex = 0
Else
CboCalType.ListIndex = -1
End If
End If
End If
End If
End Sub
Private Sub DCEmp_Name_Change()
DCEmp_Name_Click (0)
End Sub

Private Sub DCEmp_Name_Click(Area As Integer)
 Dim Equation As Double
     If val(DCEmp_Name.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DCEmp_Name.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    LBLWhereSTR.Caption = GetSpecificComponentIncalculations(val(Me.DCComponent.BoundText), Equation)
    LBLhOURrATE.Caption = Equation

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
    ChAccept.RightToLeft = False
    ChAccept.Caption = "Accept "
ISButton8.Caption = "Search"
ISButton2.Caption = "Print"
ISButton1.Caption = "Print Note"
    Me.Caption = "Punishment"
    Label1(2).Caption = Me.Caption
Label2(3).Caption = "Due Date"
    Label1(4).Caption = "Remarks"

    Label1(6).Caption = "ID"
 Label1(4).Caption = "Remark"
    Label2(0).Caption = "Date"
    Label1(0).Caption = "Employee"
    Label1(5).Caption = "Type"
    Label1(1).Caption = "Interval"
Label1(8).Caption = "Warning No"
Label1(9).Caption = " Department"
Label1(7).Caption = "Component"
Label1(12).Caption = "War.Reason"
Label1(14).Caption = "Sanction"
Label1(11).Caption = "Value"
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
    Dim Dcombos As ClsDataCombos
    Dim i As Integer
    Dim BKGrndPic As ClsBackGroundPic
    Dim My_SQL As String
    On Error GoTo ErrTrap

    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
  If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = " select id,name from mofrad where   MofrdDiscount=1"
    Else
        My_SQL = " select id,namee from mofrad where   MofrdDiscount=1"
    End If

    fill_combo DCComponent, My_SQL
    
    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboCalType
            .Clear
            .AddItem "Œ’„ √Ì«„ „‰ «·„— »"
            .AddItem "Œ’„ ﬁÌ„… ‰ﬁœÌ… „‰ «·„— »"
            End With
          Else

        With Me.CboCalType
            .Clear
            .AddItem "Days From Salry"
            .AddItem "Value"
        End With

    End If

    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset

    RsSavRec.CursorLocation = adUseClient
    My_SQL = "select * From tblKhsmEdafa where KhsmEdafa_Type=0 order by KhsmEdafa_Code"
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    'load tblUsers -----------------------------------------------
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCUser
    Me.DCUser.BoundText = user_id
     Dcombos.GetAdminSanction Me.DcbSanction
    Dcombos.GetEmpDepartments Me.DcbDept
    'load tblEmployee --------------------------------------------
    Dcombos.GetEmployees Me.DCEmp_Name
    Set cSearch = New clsDCboSearch
    Set cSearch.Client = DCEmp_Name

    For i = 2006 To 3000
        CmbYear.AddItem i
        CmbYear.ItemData(CmbYear.NewIndex) = i
    Next i
CmbMonth.Clear
    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbYear.Text = year(Date)
    CmbMonth.ListIndex = Month(Date) - 1

    DTPicker.value = Date
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

Private Sub Form_Activate()
    Me.ZOrder 0
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
Sub SaveInformationVacation(Optional ID As Integer = 0, Optional Yearr As Integer, Optional Monthh As Integer, Optional EmpID As Integer = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
str = " «·„›—œ«  «·„ €Ì—…"
Else
str = "Components Changing"
End If
sql = "select * from TblInforVacatiom where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("AbcenceID").value = ID
      Rs7("KhaseID").value = val(TxtKhsmEdafa_Code.Text)
      Rs7("EmpID").value = EmpID
      Rs7("NoDay").value = (NoDay)
      Rs7("RecordDate").value = DTPicker.value
      Rs7("RecordDateH").value = ToHijriDate(DTPicker.value)
      Rs7("TypeVacation").value = 1
      Rs7("Remarks").value = str
      Rs7("Yearr").value = Yearr
      Rs7("Monthh").value = Monthh
      Rs7.update
End Sub
Public Sub FiLLRec()
    On Error GoTo ErrTrap
Dim str As String
Dim HourRate As Double
Dim val1 As Double
Dim EmployeeSalary As Double
Dim RsDeVac As ADODB.Recordset
Dim RsDeVacDet As ADODB.Recordset
Dim StrSQL As String
Dim IDTemp As Integer
If Me.TxtModFlg.Text <> "R" Then
  StrSQL = "Delete From TblInforVacatiom Where KhaseID=" & val(TxtKhsmEdafa_Code.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where KsmID=" & val(Me.TxtKhsmEdafa_Code.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegister Where KsmID=" & val(Me.TxtKhsmEdafa_Code.Text) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
End If

    RsSavRec.Fields("KhsmEdafa_Type").value = 0
    RsSavRec.Fields("Emp_ID").value = IIf(DCEmp_Name.Text <> "", Trim(DCEmp_Name.BoundText), Null)
    RsSavRec.Fields("KhsmEdafa_Date").value = IIf(CStr(StrDate) <> "", StrDate, Null)
    RsSavRec.Fields("KhsmEdafa_Value").value = IIf(TxtKhsmEdafa_Value.Text <> "", Trim(TxtKhsmEdafa_Value.Text), Null)
    RsSavRec.Fields("KhsmEdafa_Code").value = IIf(TxtKhsmEdafa_Code.Text <> "", Trim(TxtKhsmEdafa_Code.Text), Null)
    RsSavRec.Fields("Notes").value = IIf(TxtNotes.Text <> "", Trim(TxtNotes.Text), Null)
    RsSavRec.Fields("RcDate").value = IIf(DTPicker.value <> "", Trim(DTPicker.value), Null)
    RsSavRec("CalculateValueType").value = IIf(Me.CboCalType.ListIndex = -1, 0, Me.CboCalType.ListIndex)
    RsSavRec("UserID").value = IIf(Me.DCUser.BoundText = "", user_id, Me.DCUser.BoundText)
    RsSavRec.Fields("Mofrd").value = IIf(DCComponent.Text <> "", Trim(DCComponent.BoundText), Null)
    RsSavRec.Fields("DeptID").value = IIf(DcbDept.Text <> "", Trim(DcbDept.BoundText), Null)
    RsSavRec.Fields("SanctionID").value = IIf(DcbSanction.Text <> "", Trim(DcbSanction.BoundText), Null)
    RsSavRec.Fields("Resonalarm").value = IIf(TxtResonalarm.Text <> "", Trim(TxtResonalarm.Text), Null)
    RsSavRec.Fields("Val").value = IIf(TxtVal.Text <> "", val(TxtVal.Text), Null)
    RsSavRec.Fields("AlrmOrder").value = IIf(TxtAlrmOrder.Text <> "", val(TxtAlrmOrder.Text), Null)
    RsSavRec.Fields("DateSet").value = IIf(DateSet.value <> "", Trim(DateSet.value), Null)
    If Me.ChAccept.value = vbChecked Then
     RsSavRec.Fields("Accept").value = 1
     Else
      RsSavRec.Fields("Accept").value = 0
     End If
    RsSavRec.update
''''''''''''''''''''''''Header

Set RsDeVac = New ADODB.Recordset
Set RsDeVacDet = New ADODB.Recordset
  StrSQL = "SELECT  *  from TblChangedComponentRegister Where (1 = -1)"
    RsDeVac.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
      StrSQL = "SELECT  *  from TblChangedComponentRegisterDetails Where (1 = -1)"
    RsDeVacDet.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    IDTemp = CStr(new_id("TblChangedComponentRegister", "ChangedComponentid", "", True))
     RsDeVac.AddNew
     
     RsDeVac("ChangedComponentid").value = IDTemp
     RsDeVac("KsmID").value = val(TxtKhsmEdafa_Code.Text)
     RsDeVac("Flag").value = 1
     RsDeVac("RecordDate").value = DTPicker.value
     RsDeVac("ComponentID").value = IIf(DCComponent.Text <> "", Trim(DCComponent.BoundText), Null)
     RsDeVac("year").value = IIf(CmbYear.Text <> "", val(CmbYear.ListIndex), Null)
     RsDeVac("month").value = IIf(CmbMonth.Text <> "", val(CmbMonth.ListIndex), Null)
     RsDeVac("Actualyear").value = val(CmbYear.Text)
     RsDeVac("Actualmonth").value = val(CmbMonth.ListIndex) + 1
     RsDeVac.update
     If val(CboCalType.ListIndex) = 1 Then
     EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(DCEmp_Name.BoundText), LBLWhereSTR)
    MsgBox EmployeeSalary
     End If
    If val(CboCalType.ListIndex) = 0 Then
     EmployeeSalary = GetEmployeeSalaryAccordingToComponent(val(DCEmp_Name.BoundText), LBLWhereSTR)
     
                    '«Ì«„
                    If SystemOptions.MonthIs30days = True Then
                    
                        HourRate = (EmployeeSalary / 30)
                    Else
                        HourRate = (EmployeeSalary * 12 / 365)
                    End If

                    val1 = Round(HourRate * val(LBLhOURrATE) * val(TxtKhsmEdafa_Value.Text), SystemOptions.EmpComponentDigts)
      End If
     ''''''''''''''''''/////////////////////
     RsDeVacDet.AddNew
     RsDeVacDet("ChangedComponentid").value = IDTemp
     RsDeVacDet("KsmID").value = val(TxtKhsmEdafa_Code.Text)
     RsDeVacDet("Emp_id").value = IIf(DCEmp_Name.Text <> "", Trim(DCEmp_Name.BoundText), Null)
     
     If SystemOptions.UserInterface = ArabicInterface Then
     str = "„‰ «·Œ’Ê„«  »—ﬁ„" & " " & TxtKhsmEdafa_Code.Text
     Else
     str = "From Discounts Of Number" & " " & TxtKhsmEdafa_Code.Text
     End If
                RsDeVacDet("NoofDays").value = val(TxtKhsmEdafa_Value.Text)
                RsDeVacDet("NoOfMinutes").value = 0
                RsDeVacDet("NoOfHour").value = 0
                RsDeVacDet("HourRate").value = HourRate
                RsDeVacDet("Salary").value = EmployeeSalary
                RsDeVacDet("Value").value = val1
                RsDeVacDet("remarks").value = str
             RsDeVacDet.update
             If val(CboCalType.ListIndex) = 0 Then
             SaveInformationVacation IDTemp, (val(CmbYear.ListIndex) + 2006), val(CmbMonth.ListIndex) + 1, val(Me.DCEmp_Name.BoundText), val(TxtKhsmEdafa_Value.Text)
             End If
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
Else
 MsgBox "Saved Successfully", vbOKOnly + vbMsgBoxRight, App.title
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
    
    TxtKhsmEdafa_Code.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Code").value), "", RsSavRec.Fields("KhsmEdafa_Code").value)
    TxtKhsmEdafa_Value.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Value").value), "", RsSavRec.Fields("KhsmEdafa_Value").value)
    DTPicker.value = IIf(IsNull(RsSavRec.Fields("RcDate").value), Date, RsSavRec.Fields("RcDate").value)
    TxtNotes.Text = IIf(IsNull(RsSavRec.Fields("Notes").value), "", RsSavRec.Fields("Notes").value)
    CmbMonth.ListIndex = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Date").value), -1, Month(RsSavRec.Fields("KhsmEdafa_Date").value) - 1)
    CmbYear.Text = IIf(IsNull(RsSavRec.Fields("KhsmEdafa_Date").value), "", year(RsSavRec.Fields("KhsmEdafa_Date").value))
    Me.CboCalType.ListIndex = IIf(IsNull(RsSavRec("CalculateValueType").value), 0, RsSavRec("CalculateValueType").value)
    Me.DCUser.BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)
    Me.DcbSanction.BoundText = IIf(IsNull(RsSavRec("SanctionID").value), "", RsSavRec("SanctionID").value)
    Me.DcbDept.BoundText = IIf(IsNull(RsSavRec("DeptID").value), "", RsSavRec("DeptID").value)
    Me.DCComponent.BoundText = IIf(IsNull(RsSavRec("Mofrd").value), "", RsSavRec("Mofrd").value)
    TxtResonalarm.Text = IIf(IsNull(RsSavRec.Fields("Resonalarm").value), "", RsSavRec.Fields("Resonalarm").value)
    TxtVal.Text = IIf(IsNull(RsSavRec.Fields("Val").value), "", RsSavRec.Fields("Val").value)
    TxtAlrmOrder.Text = IIf(IsNull(RsSavRec.Fields("AlrmOrder").value), "", RsSavRec.Fields("AlrmOrder").value)
    DCEmp_Name.BoundText = IIf(IsNull(RsSavRec.Fields("Emp_ID").value), "", RsSavRec.Fields("Emp_ID").value)
    DateSet.value = IIf(IsNull(RsSavRec.Fields("DateSet").value), Date, RsSavRec.Fields("DateSet").value)
    If Not (IsNull(RsSavRec.Fields("Accept").value)) Then
    If RsSavRec.Fields("Accept").value = True Then
    Me.ChAccept = vbChecked
    Else
    Me.ChAccept = vbUnchecked
    End If
    Else
     Me.ChAccept = vbUnchecked
    End If
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

ErrTrap:

End Sub
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  MySQL = " SELECT     dbo.tblKhsmEdafa.KhsmEdafa_ID, dbo.tblKhsmEdafa.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
  MySQL = MySQL + "                    dbo.tblKhsmEdafa.KhsmEdafa_Date, dbo.tblKhsmEdafa.KhsmEdafa_Type, dbo.tblKhsmEdafa.KhsmEdafa_Value, dbo.tblKhsmEdafa.KhsmEdafa_Code,"
  MySQL = MySQL + "                     dbo.tblKhsmEdafa.RcDate, dbo.tblKhsmEdafa.CalculateValueType, dbo.tblKhsmEdafa.Resonalarm, dbo.tblKhsmEdafa.Val, dbo.tblKhsmEdafa.AlrmOrder,"
  MySQL = MySQL + "                     dbo.tblKhsmEdafa.Mofrd, dbo.mofrad.name, dbo.mofrad.nameE, dbo.tblKhsmEdafa.SanctionID, dbo.TblAdminSanction.Name AS ScaName,"
  MySQL = MySQL + "                     dbo.TblAdminSanction.NameE AS ScaNameE, dbo.tblKhsmEdafa.Accept, dbo.tblKhsmEdafa.DeptID, dbo.TblEmpDepartments.DepartmentName,"
  MySQL = MySQL + "                     dbo.TblEmpDepartments.DepartmentNamee , dbo.tblKhsmEdafa.Notes, dbo.tblKhsmEdafa.DateSet "
  MySQL = MySQL + "   FROM         dbo.tblKhsmEdafa LEFT OUTER JOIN"
  MySQL = MySQL + "                     dbo.TblEmpDepartments ON dbo.tblKhsmEdafa.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  MySQL = MySQL + "                     dbo.TblAdminSanction ON dbo.tblKhsmEdafa.SanctionID = dbo.TblAdminSanction.ID LEFT OUTER JOIN"
  MySQL = MySQL + "                     dbo.mofrad ON dbo.tblKhsmEdafa.Mofrd = dbo.mofrad.id LEFT OUTER JOIN"
  MySQL = MySQL + "                     dbo.TblEmployee ON dbo.tblKhsmEdafa.Emp_ID = dbo.TblEmployee.Emp_ID"
  MySQL = MySQL + "  Where (dbo.tblKhsmEdafa.KhsmEdafa_ID =" & val(TxtKhsmEdafa_Code.Text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeductionNote.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeductionNoteE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
         xReport.ParameterFields(4).AddCurrentValue WriteNo(Me.TxtVal.Text, 1)
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
       
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Me.TxtVal.Text, 1)
        StrReportTitle = ""
   
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
     
   
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Image1_Click()

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub ISButton1_Click()
print_report
End Sub

Private Sub ISButton8_Click()
  Unload General_Search
        General_Search.send_form = "Khsm"
            Load General_Search
            General_Search.send_form = "Khsm"
            General_Search.show
             General_Search.send_form = "Khsm"
End Sub

Private Sub TxtAlrmOrder_Change()
If val(Me.TxtAlrmOrder.Text) <> 0 Then
GetEmployeeWarning val(Me.TxtAlrmOrder.Text)
End If
End Sub

Private Sub TxtAlrmOrder_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
      Unload General_Search
        General_Search.send_form = "Warning"
        General_Search.Index = 2
            Load General_Search
            General_Search.send_form = "Warning"
            General_Search.show
             General_Search.send_form = "Warning"
  
    End If
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

Private Sub TxtKhsmEdafa_Value_LostFocus()
    Dim Msg As String

    If val(TxtKhsmEdafa_Value.Text) > 90 And CboCalType.ListIndex = 0 Then

        Msg = "⁄›Ê« ·« ÌÃÊ“ Œ’„ «ﬂÀ— „‰ 90 ÌÊ„"
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtKhsmEdafa_Value.Text = ""
        TxtKhsmEdafa_Value.SetFocus
    End If

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
        '  btnNext.Enabled = False
        '  btnPrevious.Enabled = False
        '  btnFirst.Enabled = False
        '  btnLast.Enabled = False
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

    If SystemOptions.UserInterface = EnglishInterface Then
        
        If CboCalType.ListIndex = -1 Then Exit Sub
        If Me.CboCalType.ListIndex = 0 Then
            Me.Label1(1).Caption = "Interval"
            Me.Label1(3).Caption = "Day"
        ElseIf Me.CboCalType.ListIndex = 1 Then
            Me.Label1(1).Caption = "Value"
            Me.Label1(3).Caption = ""
        End If

    Else
        
        If CboCalType.ListIndex = -1 Then Exit Sub
        If Me.CboCalType.ListIndex = 0 Then
            Me.Label1(1).Caption = "„œ… «·Œ’„"
            Me.Label1(3).Caption = "ÌÊ„"
        ElseIf Me.CboCalType.ListIndex = 1 Then
            Me.Label1(1).Caption = "ﬁÌ„… «·Œ’„"
            Me.Label1(3).Caption = ""
        End If

    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg.Text <> "R" Then
   Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DCEmp_Name.BoundText = EmpID
    End If
  End If
End Sub

Private Sub TxtVal_Change()
Label1(10).Caption = WriteNo(Me.TxtVal.Text, 1)
End Sub
