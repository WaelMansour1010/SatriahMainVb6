VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCarExpireLicens 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15795
   Icon            =   "frmcarExpireLicense.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   675
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   11235
      Width           =   5115
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmcarExpireLicense.frx":058A
         Left            =   2280
         List            =   "frmcarExpireLicense.frx":059A
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   870
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   2400
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   285
         Visible         =   0   'False
         Width           =   1065
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
         Left            =   75
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·√Ã«“…"
         Top             =   285
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·ÊŸÌ›…"
         Height          =   195
         Index           =   3
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·ÊŸÌ›…"
         Height          =   285
         Index           =   0
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1890
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   -90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   10485
      Width           =   6840
      _cx             =   12065
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
         Left            =   4125
         TabIndex        =   7
         Top             =   555
         Visible         =   0   'False
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
         ButtonImage     =   "frmcarExpireLicense.frx":05B3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   2580
         TabIndex        =   8
         Top             =   555
         Visible         =   0   'False
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
         ButtonImage     =   "frmcarExpireLicense.frx":094D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   3345
         TabIndex        =   9
         Top             =   555
         Visible         =   0   'False
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
         ButtonImage     =   "frmcarExpireLicense.frx":0CE7
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   1815
         TabIndex        =   10
         Top             =   555
         Visible         =   0   'False
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
         ButtonImage     =   "frmcarExpireLicense.frx":1081
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   1050
         TabIndex        =   11
         Top             =   555
         Visible         =   0   'False
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
         ButtonImage     =   "frmcarExpireLicense.frx":141B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   12
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
         ButtonImage     =   "frmcarExpireLicense.frx":19B5
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   13
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
         ButtonImage     =   "frmcarExpireLicense.frx":1D4F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   3765
         TabIndex        =   14
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
         ButtonImage     =   "frmcarExpireLicense.frx":20E9
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
         TabIndex        =   18
         Top             =   225
         Width           =   540
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   225
         Visible         =   0   'False
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
         TabIndex        =   15
         Top             =   225
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   9210
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   15765
      _cx             =   27808
      _cy             =   16245
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
      Caption         =   "0|1"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   8790
         Left            =   45
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   45
         Width           =   15675
         _cx             =   27649
         _cy             =   15505
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
         Begin VB.Frame Frame1 
            Caption         =   "«” ⁄·«„"
            Height          =   1725
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   6675
            Width           =   12990
            Begin VB.ComboBox DcbStuts 
               Height          =   315
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   930
               Width           =   1440
            End
            Begin VB.TextBox TxtQtyUpload 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   4020
               TabIndex        =   64
               Top             =   1230
               Width           =   1290
            End
            Begin VB.CommandButton Command1 
               Caption         =   "«” ⁄·«„"
               Height          =   315
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command2 
               Caption         =   "ÿ»«⁄Â"
               Height          =   315
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   600
               Width           =   1095
            End
            Begin Dynamic_Byte.NourHijriCal Txt_to_H 
               Height          =   255
               Left            =   4920
               TabIndex        =   33
               Top             =   600
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
            End
            Begin Dynamic_Byte.NourHijriCal Txt_from_H 
               Height          =   255
               Left            =   4920
               TabIndex        =   34
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
            End
            Begin MSComCtl2.DTPicker d1 
               Height          =   315
               Left            =   3000
               TabIndex        =   35
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   155779073
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker d2 
               Height          =   315
               Left            =   3000
               TabIndex        =   36
               Top             =   600
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Format          =   155779073
               CurrentDate     =   38784
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   0
               Left            =   3510
               TabIndex        =   59
               Top             =   1230
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   ">"
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   60
               Top             =   1260
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "<"
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   61
               Top             =   1260
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "="
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   3
               Left            =   1560
               TabIndex        =   62
               Top             =   1260
               Width           =   615
               _Version        =   786432
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   ">="
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   4
               Left            =   600
               TabIndex        =   63
               Top             =   1260
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "<="
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ«·…«·„⁄œ…"
               Height          =   210
               Index           =   47
               Left            =   4575
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   960
               Width           =   795
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   315
               Index           =   7
               Left            =   5355
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   1230
               Width           =   1065
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "„‰"
               Height          =   255
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "«·Ï"
               Height          =   255
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   600
               Width           =   255
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   0
            Width           =   15675
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   26
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
                  TabIndex        =   27
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
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Text            =   "modflag"
               Top             =   120
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
               TabIndex        =   23
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   480
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
                     Picture         =   "frmcarExpireLicense.frx":2483
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":281D
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":2BB7
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":2F51
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":32EB
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":3685
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":3A1F
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmcarExpireLicense.frx":3FB9
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„⁄œ« /«·”Ì«—«  «· Ì ” ‰ ÂÌ «” „«—« Â«"
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
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   120
               Width           =   3735
            End
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   465
            Left            =   345
            TabIndex        =   21
            Top             =   8025
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   820
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
            ButtonImage     =   "frmcarExpireLicense.frx":4353
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   5640
            Left            =   0
            TabIndex        =   29
            Top             =   840
            Width           =   15585
            _cx             =   27490
            _cy             =   9948
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmcarExpireLicense.frx":46ED
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   8790
         Left            =   16410
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   45
         Width           =   15675
         _cx             =   27649
         _cy             =   15505
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
         Begin VB.TextBox TxtQtyDownLoad 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   10350
            TabIndex        =   72
            Top             =   1635
            Width           =   1290
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1650
            Width           =   3495
            Begin XtremeSuiteControls.RadioButton RdNet 
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   67
               Top             =   0
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   ">"
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdNet 
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   68
               Top             =   0
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "<"
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdNet 
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   69
               Top             =   0
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "="
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdNet 
               Height          =   255
               Index           =   3
               Left            =   1080
               TabIndex        =   70
               Top             =   0
               Width           =   615
               _Version        =   786432
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   ">="
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdNet 
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   71
               Top             =   0
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "<="
               BackColor       =   14871017
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
         End
         Begin VB.Frame FRMCAncel 
            Caption         =   "Õœœ «·”»» ·«Ìﬁ«› «·ŒÿÂ"
            Height          =   4335
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3900
            Visible         =   0   'False
            Width           =   7695
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ìﬁ«› ﬂ«„· «·Œÿ…"
               Height          =   735
               Index           =   0
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   2040
               Width           =   3135
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "«Ìﬁ«› «·”ÿ— «·„Õœœ ›ﬁÿ  "
               Height          =   615
               Index           =   1
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   2520
               Value           =   -1  'True
               Width           =   3135
            End
            Begin VB.TextBox CancelReason 
               Alignment       =   1  'Right Justify
               Height          =   1215
               Left            =   810
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   56
               Top             =   510
               Width           =   6135
            End
            Begin VB.CommandButton Command3 
               Caption         =   " √ﬂÌœ «·«Ìﬁ«›"
               Height          =   495
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   3480
               Width           =   2775
            End
         End
         Begin VB.TextBox TxtID 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   270
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1095
            Width           =   1665
         End
         Begin VB.CheckBox chkDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì  «—ÌŒ „Õœœ"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   13530
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1095
            Width           =   1515
         End
         Begin C1SizerLibCtl.C1Elastic EleHeader 
            Height          =   990
            Left            =   0
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   15645
            _cx             =   27596
            _cy             =   1746
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   24
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
            BackColor       =   16777215
            ForeColor       =   4210688
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   0
            ChildSpacing    =   0
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰»ÌÂ«  ’Ì«‰… «·„⁄œ« /«·”Ì«—«  / «·„⁄œ«  "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   30
               Left            =   8760
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   120
               Width           =   4770
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid13 
            Height          =   5880
            Left            =   180
            TabIndex        =   42
            Top             =   2145
            Width           =   15315
            _cx             =   27014
            _cy             =   10372
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
            BackColorAlternate=   16777088
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmcarExpireLicense.frx":48DB
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
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   615
            Left            =   345
            TabIndex        =   43
            Top             =   7905
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   1085
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
            ButtonImage     =   "frmcarExpireLicense.frx":4AD5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   615
            Left            =   4440
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7905
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1085
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
            ButtonImage     =   "frmcarExpireLicense.frx":4E6F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   615
            Left            =   2415
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   7905
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1085
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
            ButtonImage     =   "frmcarExpireLicense.frx":5209
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSComCtl2.DTPicker fromDate 
            Height          =   405
            Left            =   10785
            TabIndex        =   46
            Top             =   1095
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   157155329
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   405
            Left            =   7875
            TabIndex        =   47
            Top             =   1095
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   157155329
            CurrentDate     =   41640
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   405
            Left            =   3315
            TabIndex        =   52
            Top             =   1095
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   " ‘«‘… »ÕÀ Œÿ… «·’Ì«‰Â"
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
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "frmcarExpireLicense.frx":BA6B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«Ì«„"
            Height          =   315
            Index           =   0
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1635
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ŒÿÂ"
            Height          =   375
            Index           =   5
            Left            =   2055
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1095
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   375
            Index           =   4
            Left            =   9540
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   1110
            Width           =   660
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   375
            Index           =   1
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   1110
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "FrmCarExpireLicens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim My_SQL As String
Dim date_type As Integer
Dim xApp As New CRAXDRT.Application
Dim EmpReport As ClsEmployeeReport
Dim Askinterval As String
Dim Askcount As Integer
Public Indx As Integer
Public SQLG As String


Private Sub TxtQtyUpload_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtQtyUpload.Text, 0)
End Sub

Private Sub ALLButton1_Click()
'  FrmSearchCarsPlan.Indx = 1
'  Load FrmSearchCarsPlan
'  FrmSearchCarsPlan.Indx = 1
'  FrmSearchCarsPlan.show vbModal
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

    If TxtVac_ID.Text <> "" Then
        If CheckDelJobType(val(Me.TxtVac_ID.Text)) = False Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
            RsSavRec.Find "JobTypeID=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
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
        Me.TxtVacName.SetFocus
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
 
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"

    My_SQL = "TblEmpJobsTypes"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
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

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblEmpJobsTypes", "JobTypeName", Trim(TxtVacName.Text), "JobTypeName", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
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

Private Sub chkDate_Click()
    If chkDate.value = vbChecked Then
        FromDate.Enabled = True
        toDate.Enabled = True
        FromDate.value = Date
        toDate.value = Date
        
    ElseIf chkDate.value = vbUnchecked Then
        FromDate.Enabled = False
        toDate.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrTrap
 
    If date_type = 1 Then
        d1.value = Format$(ToGregorianDate(Txt_from_H.value), "dd-mm-yyyy")
        d2.value = Format$(ToGregorianDate(Txt_to_H.value), "dd-mm-yyyy")
    End If

    My_SQL = "SELECT   * from dbo.TblCarsData WHERE  1 = 1  "
    If Trim(DcbStuts.Text) <> "" Then
        My_SQL = My_SQL & " AND  StutsID = " & val(DcbStuts.ListIndex) & ""
    End If
    If val(TxtQtyUpload.Text) <> 0 Then
        If RdTotal(0).value = True Then
            My_SQL = My_SQL & " AND  DATEDiff(D ,GETDATE(), TblCarsData.LicenseExpireDate)  >" & val(TxtQtyUpload.Text) & ""
        ElseIf RdTotal(1).value = True Then
            My_SQL = My_SQL & " AND  DATEDiff(D ,GETDATE(), TblCarsData.LicenseExpireDate)  <" & val(TxtQtyUpload.Text) & ""
        ElseIf RdTotal(2).value = True Then
            My_SQL = My_SQL & " AND  DATEDiff(D ,GETDATE(), TblCarsData.LicenseExpireDate)  =" & val(TxtQtyUpload.Text) & ""
        ElseIf RdTotal(3).value = True Then
            My_SQL = My_SQL & " AND  DATEDiff(D ,GETDATE(), TblCarsData.LicenseExpireDate)  >=" & val(TxtQtyUpload.Text) & ""
        ElseIf RdTotal(4).value = True Then
            My_SQL = My_SQL & " AND  DATEDiff(D ,GETDATE(), TblCarsData.LicenseExpireDate)  <=" & val(TxtQtyUpload.Text) & ""
        End If
    Else
'        My_SQL = My_SQL & "  and (LicenseExpireDate >= CONVERT(DATETIME, '" & Format$(d1.value, "dd-mm-yyyy") & " 00:00:00', 102)) "
'        My_SQL = My_SQL & "  AND (LicenseExpireDate <= CONVERT(DATETIME, '" & Format$(d2.value, "dd-mm-yyyy") & " 00:00:00', 102))"
        
        My_SQL = My_SQL & "  and (LicenseExpireDate >=  " & SQLDate(d1.value, True) & ")"
        My_SQL = My_SQL & "  and (LicenseExpireDate >=  " & SQLDate(d2.value, True) & ")"
        
 
    End If


    
    FillGridWithData

    'My_SQL = "select * From TblEmployee  where    (MONTH(DateEndekama) <= MONTH(GETDATE()))"
    'End If

    Exit Sub
ErrTrap:
    MsgBox "«œŒ·   «—ÌŒ ÂÃ—Ì Œ«ÿÌ¡", vbCritical
End Sub

Private Sub Command2_Click()
    Dim rs As New ADODB.Recordset
 
    Dim xReport As New CRAXDRT.Report

    '    Sql = "SELECT * from emp_all_details WHERE emp_code='" & FrmEmployee.TxtEmp_Code.text & "'"
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(App.path & "\reports\Transporter\REPORT1.rpt")
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = (App.path & "\reports\Transporter\REPORT1.rpt")
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    xReport.ParameterFields(1).AddCurrentValue Txt_from_H.value
    xReport.ParameterFields(2).AddCurrentValue Txt_to_H.value
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    SendKeys "{RIGHT}"
      
End Sub

Private Sub Command3_Click()
Dim StrSQL As String
Dim palnID As Double
Dim MaintenanceTypesLineNo As Double
        If CancelReason = "" Then
                    MsgBox "«œŒ· «·”»» «Ê·«", vbCritical
                    Exit Sub
        End If

   palnID = val(Me.VSFlexGrid13.TextMatrix(VSFlexGrid13.Row, Me.VSFlexGrid13.ColIndex("PlanCode")))
        MaintenanceTypesLineNo = val(Me.VSFlexGrid13.TextMatrix(VSFlexGrid13.Row, Me.VSFlexGrid13.ColIndex("MaintenanceTypesLineNo")))
 
         If Option1(0).value = True Then ' ﬂ· «·ŒÿÂ
         StrSQL = " update  TblCarMaintenancePlanDetails set done =2  ,  cancelreason ='" & Me.CancelReason & "'where planid=" & palnID
         Else '«·„Õœœ ›ﬁÿ
         StrSQL = " update  TblCarMaintenancePlanDetails set done =2  ,  cancelreason ='" & Me.CancelReason & "'where id=" & MaintenanceTypesLineNo
         End If
        Cn.Execute StrSQL
          FRMCAncel.Visible = False
          FillGridWithData2
End Sub

Private Sub d1_Change()
    date_type = 2
End Sub

Private Sub d2_Change()
    date_type = 2
End Sub

Private Sub Form_Load()
    d1.value = Date
    d2.value = Date
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ' On Error GoTo ErrTrap
    Dim i As Integer
 
    My_SQL = "TblCarsData"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient

    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireLicense", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "Count_ExpireLicense", 0)
    My_SQL = "SELECT     * from dbo.TblCarsData Where LicenseExpireDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"

    'My_SQL = "SELECT     * from dbo.TblEmployee Where (Month(DateEndekama) <= Month(GetDate())) And (year(DateEndekama) <= year(GetDate()))"

    With Me.Grid

        '    .Cell(flexcpPicture, 0, .ColIndex("Emp_Name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        '    .Cell(flexcpPicture, 0, .ColIndex("DateEndPasp")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        '    .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    C1Tab1.TabVisible(0) = False
    C1Tab1.TabVisible(1) = False

  
    C1Tab1.TabVisible(Indx) = True
    C1Tab1.CurrTab = Indx
    If Indx = 0 Then
        FillGridWithData
    ElseIf Indx = 1 Then
        FillGridWithData2
    End If
    
    

    If SystemOptions.UserInterface = ArabicInterface Then
        With DcbStuts
            .Clear
            .AddItem "‰‘ÿ"
            .AddItem " Õ  «·’Ì«‰…"
            .AddItem "„»«⁄"
        End With
    Else
        With DcbStuts
            .Clear
            .AddItem "Active"
            .AddItem "Under Maintenance "
            .AddItem "Sold"
        End With
    End If
    
    
    
    'BtnFirst_Click
    ShowTip

ErrTrap:
End Sub

Function ChangeLang()
    Me.Caption = "Expire Residence"
    Label1(2).Caption = Me.Caption
    Frame1.Caption = "Query"
    Label3.Caption = "From"
    Label4.Caption = "To"
    Command1.Caption = "Search"
    Command2.Caption = "Print"
    btnCancel.Caption = "Exit"



    Label1(30).Caption = "Equipment Maintenance Alarm"
    ISButton2.Caption = "Update"
    ISButton3.Caption = "Print"
    ISButton1.Caption = "Exit"
    With Me.VSFlexGrid13
        .TextMatrix(0, .ColIndex("Convert")) = "Convert"
        .TextMatrix(0, .ColIndex("CarCode")) = "Equipment code"
        .TextMatrix(0, .ColIndex("PlanCode")) = "Plan Code"
        .TextMatrix(0, .ColIndex("CarName")) = "Equipment Name"
        .TextMatrix(0, .ColIndex("AlaramDate")) = "Alarm Date"
        .TextMatrix(0, .ColIndex("MantType")) = "Maintenance Type"
    End With
    
    With Me.Grid
        '.TextMatrix(0, .ColIndex("emp_code")) = "Emp Code"
        '.TextMatrix(0, .ColIndex("emp_name")) = "Emp Name"
        '.TextMatrix(0, .ColIndex("NumEkama")) = "Num."
        '.TextMatrix(0, .ColIndex("placeEkama")) = "Issue Place"
        '.TextMatrix(0, .ColIndex("DateExpoekamaH")) = "Issue Date H"
        '.TextMatrix(0, .ColIndex("DateEndekamah")) = "Expire Date H"
        '.TextMatrix(0, .ColIndex("DateExpoekama")) = "Issue Date G"
        '.TextMatrix(0, .ColIndex("DateEndekama")) = "Expire Date G"
        '.TextMatrix(0, .ColIndex("days")) = "Remain Days"
        'FillGridWithData
    End With

End Function

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

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblEmpJobsTypes", "JobTypeID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("JobTypeID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("JobTypeName").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("JobTypeID").value), "", RsSavRec.Fields("JobTypeID").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("JobTypeName").value), "", RsSavRec.Fields("JobTypeName").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("JobTypeID")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
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


Private Sub FromDate_Change()
    FillGridWithData2
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("JobTypeID")))
ErrTrap:
End Sub

Private Sub ISButton1_Click()
    Unload Me
End Sub

Private Sub ISButton2_Click()
    FillGridWithData2
End Sub

Private Sub ISButton3_Click()
    print_report
End Sub
Private Sub ToDate_Change()
    FillGridWithData2
End Sub

Private Sub Txt_from_H_GotFocus()
    date_type = 1
End Sub

Private Sub Txt_to_H_GotFocus()
    date_type = 1
End Sub

Private Sub TxtID_Change()
FillGridWithData2
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "JobTypeID=" & RecId, , adSearchForward, 1

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
    
        '    btnNext.Enabled = True
        '    btnPrevious.Enabled = True
        '    btnFirst.Enabled = True
        '    btnLast.Enabled = True
    
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
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1

                .TextMatrix(i, .ColIndex("Ser")) = i
                '
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)

                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
                'If FrmEmployee.OptExpirEkama = True Then
     
                .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs.Fields("BoardNO").value), "", rs.Fields("BoardNO").value)
              
                .TextMatrix(i, .ColIndex("LicenseExpireDateH")) = IIf(IsNull(rs.Fields("LicenseExpireDateH").value), "", rs.Fields("LicenseExpireDateH").value)
                .TextMatrix(i, .ColIndex("LicenseExpireDate")) = IIf(IsNull(rs.Fields("LicenseExpireDate").value), "", rs.Fields("LicenseExpireDate").value)
            
                Dim days As Double
      
                days = IIf(IsNull(rs.Fields("LicenseExpireDate").value), "", DateDiff("d", Date, rs.Fields("LicenseExpireDate").value))
           
                If days >= 0 Then
           
                    .TextMatrix(i, .ColIndex("days")) = days
           
                Else
           
                    .TextMatrix(i, .ColIndex("LateDate")) = Abs(days)
              
                End If
           
                '         .TextMatrix(i, .ColIndex("days")) = IIf(IsNull(rs.Fields("LicenseExpireDate").value), _
                          "", DateDiff("d", Date, rs.Fields("LicenseExpireDate").value))
           
                If days = 0 Then
                    .Cell(flexcpBackColor, i, 1, i, 10) = vbYellow
                ElseIf days <= 0 Then
                    .Cell(flexcpBackColor, i, 1, i, 10) = vbRed
              
                End If
 
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
        '    .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        '    .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        '    .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        '    .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        '    .AddControl btnLast, Msg, True
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
    'If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
    '    If btnFirst.Enabled = False Then Exit Sub
    '    BtnFirst_Click
    'End If
    'Move Previous---------------------------------------------------------
    'If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
    '    If btnPrevious.Enabled = False Then Exit Sub
    '    BtnPrevious_Click
    'End If

    'Move Next---------------------------------------------------------
    'If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
    '    If btnNext.Enabled = False Then Exit Sub
    '    BtnNext_Click
    'End If

    'Move Last---------------------------------------------------------
    'If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
    '    If btnLast.Enabled = False Then Exit Sub
    '    BtnLast_Click
    'End If

    'End If

    Exit Sub
ErrTrap:
End Sub

Private Function CheckDelJobType(LngJobTypeID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where JobTypeID=" & LngJobTypeID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelJobType = False
    Else
        CheckDelJobType = True
    End If

    rs.Close
    Set rs = Nothing
End Function
Public Sub FillGridWithData2()
    'On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    
    SQLG = "SELECT    TblCarMaintenancePlanDetails.Id MaintenanceTypesLineNo, TblCarMaintenancePlanDetails.alarmType, TblCarMaintenancePlanDetails.Done, TblCarMaintenancePlanDetails.DoneDate, TblCarMaintenancePlanDetails.UserID, TblCarMaintenancePlanDetails.CurrentKM, "
    SQLG = SQLG + " TblCarMaintenancePlanDetails.AlarmInKM, TblCarMaintenancePlanDetails.AlarmINDate, TblCarMaintenancePlanDetails.AlarmINTime, TblCarMaintenancePlanDetails.NoHour, TblCarMaintenancePlanDetails.GroupID,"
    SQLG = SQLG + " TblCarMaintenancePlan.RecordDate,TblCarMaintenancePlan.Planid, TblCarsData.Fullcode, TblMaintenanceType.name AS mainType, TblMaintenanceType.namee AS mainTypee, TblCarsData.Name AS carName, TblCarMaintenancePlan.RecordDate,"
    SQLG = SQLG + " TblCarsData.EqupName"
    SQLG = SQLG + " FROM TblCarsData RIGHT OUTER JOIN"
    SQLG = SQLG + " TblCarMaintenancePlan ON TblCarsData.id = TblCarMaintenancePlan.CarId FULL OUTER JOIN"
    SQLG = SQLG + " TblCarMaintenancePlanDetails LEFT OUTER JOIN"
    SQLG = SQLG + " TblMaintenanceType ON TblCarMaintenancePlanDetails.MaintenanceID = TblMaintenanceType.id ON TblCarMaintenancePlan.Planid = TblCarMaintenancePlanDetails.Planid"
  'salim here new
  SQLG = " SELECT     dbo.TblCarMaintenancePlanDetails.id AS MaintenanceTypesLineNo, dbo.TblCarMaintenancePlanDetails.alarmType, dbo.TblCarMaintenancePlanDetails.Done,"
   SQLG = SQLG + "                   dbo.TblCarMaintenancePlanDetails.DoneDate, dbo.TblCarMaintenancePlanDetails.UserID, dbo.TblCarMaintenancePlanDetails.CurrentKM,"
SQLG = SQLG + "                      dbo.TblCarMaintenancePlanDetails.AlarmInKM, dbo.TblCarMaintenancePlanDetails.AlarmINDate, dbo.TblCarMaintenancePlanDetails.AlarmINTime,"
SQLG = SQLG + "                      dbo.TblCarMaintenancePlanDetails.NoHour, dbo.TblCarMaintenancePlanDetails.GroupID, dbo.TblCarMaintenancePlan.RecordDate,"
SQLG = SQLG + "                      dbo.TblCarMaintenancePlan.Planid, dbo.TblCarsData.Fullcode, dbo.TblMaintenanceType.name AS mainType, dbo.TblMaintenanceType.namee AS mainTypee,"
SQLG = SQLG + "                      dbo.TblCarsData.Name AS carName, dbo.TblCarsData.EqupName, dbo.TblCarsData.Emp_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
SQLG = SQLG + "                      dbo.TblEmployee.Fullcode AS EmpFullCode"
SQLG = SQLG + " FROM         dbo.TblCarsData INNER JOIN"
SQLG = SQLG + "                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID RIGHT OUTER JOIN"
SQLG = SQLG + "                      dbo.TblCarMaintenancePlan ON dbo.TblCarsData.id = dbo.TblCarMaintenancePlan.CarId FULL OUTER JOIN"
SQLG = SQLG + "                      dbo.TblCarMaintenancePlanDetails LEFT OUTER JOIN"
SQLG = SQLG + "                      dbo.TblMaintenanceType ON dbo.TblCarMaintenancePlanDetails.MaintenanceID = dbo.TblMaintenanceType.id ON"
SQLG = SQLG + "                      dbo.TblCarMaintenancePlan.PlanID = dbo.TblCarMaintenancePlanDetails.PlanID"
    
    SQLG = SQLG + " Where (TblCarMaintenancePlanDetails.done = 0 Or TblCarMaintenancePlanDetails.done Is Null)"
    
   ' If chkDate.value = vbChecked Then
   '     If Not IsNull(Me.Fromdate.value) And Not IsNull(Me.todate.value) Then
   '         If Not IsNull(Me.Fromdate.value) Then
   '             SQLG = SQLG + " and AlarmINDate >=" & SQLDate(Me.Fromdate.value, True) & ""
   '         End If
  '
  '          If Not IsNull(Me.todate.value) Then
  '              SQLG = SQLG + " and AlarmINDate <=" & SQLDate(Me.todate.value, True) & ""
  '          End If
  '      Else
  '          SQLG = SQLG + " and AlarmINDate <= " & SQLDate(Date, True)
  '      End If
  '  Else
  '      SQLG = SQLG + " and AlarmINDate <= " & SQLDate(Date, True)
  '  End If
  
  
  
    If date_type = 1 Then
        d1.value = Format$(ToGregorianDate(Txt_from_H.value), "dd-mm-yyyy")
        d2.value = Format$(ToGregorianDate(Txt_to_H.value), "dd-mm-yyyy")
    End If

    My_SQL = "SELECT   * from dbo.TblCarsData WHERE  1 = 1  "
    
    If val(TxtQtyDownLoad.Text) <> 0 Then
        If RdNet(0).value = True Then
            SQLG = SQLG & " AND  DATEDiff(D ,GETDATE(), RecordDate)  >" & val(TxtQtyDownLoad.Text) & ""
        ElseIf RdNet(1).value = True Then
            SQLG = SQLG & " AND  DATEDiff(D ,GETDATE(), RecordDate)  <" & val(TxtQtyDownLoad.Text) & ""
        ElseIf RdNet(2).value = True Then
            SQLG = SQLG & " AND  DATEDiff(D ,GETDATE(), RecordDate)  =" & val(TxtQtyDownLoad.Text) & ""
        ElseIf RdNet(3).value = True Then
            SQLG = SQLG & " AND  DATEDiff(D ,GETDATE(), RecordDate)  >=" & val(TxtQtyDownLoad.Text) & ""
        ElseIf RdNet(4).value = True Then
            SQLG = SQLG & " AND  DATEDiff(D ,GETDATE(), RecordDate)  <=" & val(TxtQtyDownLoad.Text) & ""
        End If
    Else
            If chkDate.value = vbChecked Then
        If Not IsNull(Me.FromDate.value) And Not IsNull(Me.toDate.value) Then
            If Not IsNull(Me.FromDate.value) Then
                SQLG = SQLG + " and RecordDate >=" & SQLDate(Me.FromDate.value, True) & ""
            End If
  
            If Not IsNull(Me.toDate.value) Then
                SQLG = SQLG + " and RecordDate <=" & SQLDate(Me.toDate.value, True) & ""
            End If
        Else
            SQLG = SQLG + " and RecordDate <= " & SQLDate(Date, True)
        End If
    Else
        SQLG = SQLG + " and RecordDate <= " & SQLDate(Date, True)
    End If


    End If


    
 
  
    
If val(Me.txtid.Text) <> 0 Then
SQLG = SQLG + " and  TblCarMaintenancePlan.PlanID=" & val(Me.txtid.Text) & ""
End If

 SQLG = SQLG & " order by RecordDate,TblCarMaintenancePlan.planid "
    Set rs = New ADODB.Recordset
    
    rs.Open SQLG, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.VSFlexGrid13
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                
                 
               .TextMatrix(i, .ColIndex("MaintenanceTypesLineNo")) = IIf(IsNull(rs.Fields("MaintenanceTypesLineNo").value), "", rs.Fields("MaintenanceTypesLineNo").value)
               

                .TextMatrix(i, .ColIndex("CarCode")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)
                .TextMatrix(i, .ColIndex("PlanCode")) = IIf(IsNull(rs.Fields("Planid").value), "", rs.Fields("Planid").value)
                If rs.Fields("carName").value = "" Or IsNull(rs.Fields("carName").value) Then
                    .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(rs.Fields("EqupName").value), "", rs.Fields("EqupName").value)
                Else
                    .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(rs.Fields("carName").value), "", rs.Fields("carName").value)
                End If
                
                
                .TextMatrix(i, .ColIndex("AlaramDate")) = IIf(IsNull(rs.Fields("AlarmINDate").value), "", rs.Fields("AlarmINDate").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(i, .ColIndex("Drivername")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                    .TextMatrix(i, .ColIndex("MantType")) = IIf(IsNull(rs.Fields("mainType").value), "", rs.Fields("mainType").value)
                Else
                .TextMatrix(i, .ColIndex("Drivername")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
                    .TextMatrix(i, .ColIndex("MantType")) = IIf(IsNull(rs.Fields("mainTypee").value), "", rs.Fields("mainTypee").value)
                End If
              '  .TextMatrix(i, .ColIndex("1")) = " ÕÊÌ· «·Ï «„— ‘€·"
                rs.MoveNext
            Next i
        End If
    End With

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
 
 
    MySQL = "SELECT TblCarMaintenancePlanDetails.alarmType, TblCarMaintenancePlanDetails.Done, TblCarMaintenancePlanDetails.DoneDate, TblCarMaintenancePlanDetails.UserID, TblCarMaintenancePlanDetails.CurrentKM, "
    MySQL = MySQL + " TblCarMaintenancePlanDetails.AlarmInKM, TblCarMaintenancePlanDetails.AlarmINDate, TblCarMaintenancePlanDetails.AlarmINTime, TblCarMaintenancePlanDetails.NoHour, TblCarMaintenancePlanDetails.GroupID,"
    MySQL = MySQL + " TblCarMaintenancePlan.Planid, TblCarsData.Fullcode, TblMaintenanceType.name AS mainType, TblMaintenanceType.namee AS mainTypee, TblCarsData.Name AS carName, TblCarMaintenancePlan.RecordDate,"
    MySQL = MySQL + " TblCarsData.EqupName"
    MySQL = MySQL + " FROM TblCarsData RIGHT OUTER JOIN"
    MySQL = MySQL + " TblCarMaintenancePlan ON TblCarsData.id = TblCarMaintenancePlan.CarId FULL OUTER JOIN"
    MySQL = MySQL + " TblCarMaintenancePlanDetails LEFT OUTER JOIN"
    MySQL = MySQL + " TblMaintenanceType ON TblCarMaintenancePlanDetails.MaintenanceID = TblMaintenanceType.id ON TblCarMaintenancePlan.Planid = TblCarMaintenancePlanDetails.Planid"
    MySQL = MySQL + " Where (TblCarMaintenancePlanDetails.done = 0 Or TblCarMaintenancePlanDetails.done Is Null)"

'SQLG = SQLG & " order by RecordDate,planid "
    If chkDate.value = vbChecked Then
        If Not IsNull(Me.FromDate.value) And Not IsNull(Me.toDate.value) Then
            If Not IsNull(Me.FromDate.value) Then
                MySQL = MySQL + " and AlarmINDate >=" & SQLDate(Me.FromDate.value, True) & ""
            End If
  
            If Not IsNull(Me.toDate.value) Then
                MySQL = MySQL + " and AlarmINDate <=" & SQLDate(Me.toDate.value, True) & ""
            End If
        Else
            MySQL = MySQL + " and AlarmINDate <= " & SQLDate(Date, True)
        End If
    Else
        MySQL = MySQL + " and AlarmINDate <= " & SQLDate(Date, True)
    End If
    MySQL = SQLG
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarMintAlarm.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarMintAlarmE.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
       
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
      '  StrReportTitle = "" '& StrAccountName

    Else
 '
 '       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
      '  StrReportTitle = ""
    End If
 xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub VSFlexGrid13_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
 If Row = 0 Then Exit Sub
 Select Case VSFlexGrid13.ColKey(Col)
 
 Case "CancelPlan"
 FRMCAncel.Visible = True
 
 Case "Show"
 FrmCarsPlan.show
 FrmCarsPlan.Retrive val(Me.VSFlexGrid13.TextMatrix(Row, Me.VSFlexGrid13.ColIndex("PlanCode")))
 
 Case "CancelPlan"
 
 
         
    Case "Convert"
        If SystemOptions.AllowConvertAlertToJob Then
            Dim Msg As String, IntRes As Integer
    
    
            Msg = "Â· «‰  „ √ﬂœ „‰  ÕÊÌ· «·Œÿ… «·Ï «„— ‘€· ..ø"
            IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
    
            If IntRes = vbYes Then
                 Load FrmOrderMaintin
                 FrmOrderMaintin.Cmd_Click (0)
                 FrmOrderMaintin.BaisedOn(1).value = True
                 FrmOrderMaintin.BaisedOn_Click (1)
                 
                 FrmOrderMaintin.MaintPlan.Text = val(Me.VSFlexGrid13.TextMatrix(Row, Me.VSFlexGrid13.ColIndex("PlanCode")))
                 FrmOrderMaintin.DCMaintenanceTypes.BoundText = val(Me.VSFlexGrid13.TextMatrix(Row, Me.VSFlexGrid13.ColIndex("MaintenanceTypesLineNo")))
                 
                 FrmOrderMaintin.DcbType.ListIndex = 0
                 FrmOrderMaintin.Frame4.Visible = True
                 FrmOrderMaintin.show
                 
            End If
        End If
    End Select
        
End Sub

Private Sub VSFlexGrid13_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If VSFlexGrid13.ColIndex("Convert") = Col Then
    VSFlexGrid13.EditMaxLength = 10
ElseIf VSFlexGrid13.ColIndex("Show") = Col Then
    VSFlexGrid13.EditMaxLength = 10
    ElseIf VSFlexGrid13.ColIndex("CancelPlan") = Col Then
    Me.FRMCAncel.Visible = True
Else
    Cancel = True
End If
End Sub
