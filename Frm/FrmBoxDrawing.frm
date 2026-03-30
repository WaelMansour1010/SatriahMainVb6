VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmBoxDrawing 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ÕÊÌ·«  „«·ÌÂ"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14745
   Icon            =   "FrmBoxDrawing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   57
      Top             =   1920
      Width           =   14775
      _cx             =   26061
      _cy             =   9128
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
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   14871017
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ«·… «·«⁄ „«œ"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4800
         Left            =   45
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   45
         Width           =   14685
         _cx             =   25903
         _cy             =   8467
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕÊÌ· »Ì‰ «·›—Ê⁄"
            Height          =   3495
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   120
            Width           =   14775
            Begin VB.ComboBox CboPaymentType 
               Height          =   315
               ItemData        =   "FrmBoxDrawing.frx":058A
               Left            =   10560
               List            =   "FrmBoxDrawing.frx":058C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   600
               Width           =   2775
            End
            Begin VB.Frame FraNote 
               BackColor       =   &H00E2E9E9&
               Height          =   1845
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   990
               Width           =   4155
               Begin VB.TextBox TxtChequeNumber1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   840
                  Width           =   2685
               End
               Begin MSComCtl2.DTPicker DtpChequeDueDate1 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   119
                  Top             =   1140
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   241631233
                  CurrentDate     =   39614
               End
               Begin MSDataListLib.DataCombo Dcbank1 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   120
                  Top             =   480
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox1 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   121
                  Top             =   120
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcAccounts1 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   122
                  Top             =   1440
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   285
                  Index           =   21
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   180
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   285
                  Index           =   22
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   510
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·‘Ìﬂ/ÕÊ«·Â"
                  Height          =   285
                  Index           =   23
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
                  Height          =   285
                  Index           =   24
                  Left            =   2820
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   1140
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Õ”«»"
                  Height          =   345
                  Index           =   17
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   1440
                  Width           =   1305
               End
            End
            Begin VB.ComboBox CboPaymentType1 
               Height          =   315
               Left            =   6000
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   116
               Top             =   600
               Width           =   2775
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   1365
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   960
               Width           =   4155
               Begin MSDataListLib.DataCombo Dcbank2 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   110
                  Top             =   480
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox2 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   111
                  Top             =   120
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCAccounts2 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   112
                  Top             =   960
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   285
                  Index           =   35
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   510
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   285
                  Index           =   36
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   180
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Õ”«»"
                  Height          =   345
                  Index           =   18
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   960
                  Width           =   1305
               End
            End
            Begin VB.TextBox TxtPerson2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   2520
               Width           =   2685
            End
            Begin VB.Frame Frame6 
               Caption         =   "„’«—Ì› ÕÊ«·… »‰ﬂÌ…"
               Height          =   615
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   2880
               Width           =   9015
               Begin VB.TextBox txtTransferExpensesBranch 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   6600
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   240
                  Width           =   855
               End
               Begin VB.CheckBox ChkToCommision 
                  Alignment       =   1  'Right Justify
                  Caption         =   " Õ„· ⁄·Ì «·„ÕÊ· «·ÌÂ"
                  CausesValidation=   0   'False
                  Height          =   255
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—”Ê„ «·ÕÊ«·Â"
                  Height          =   255
                  Left            =   7680
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin MSDataListLib.DataCombo DcBranch1 
               Height          =   315
               Left            =   10560
               TabIndex        =   129
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcBranch2 
               Height          =   315
               Left            =   6000
               TabIndex        =   130
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄ «·„ÕÊ· «·ÌÂ"
               Height          =   285
               Index           =   15
               Left            =   8820
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   285
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·›—⁄ «·„ÕÊ· „‰Â"
               Height          =   225
               Index           =   16
               Left            =   13380
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   195
               Index           =   25
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   630
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·«” ·«„"
               Height          =   195
               Index           =   28
               Left            =   8820
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   630
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” ›Ìœ"
               Height          =   285
               Index           =   26
               Left            =   8640
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   2520
               Width           =   1395
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   $"FrmBoxDrawing.frx":058E
               Height          =   2655
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   720
               Width           =   5055
            End
            Begin VB.Shape Shape1 
               BorderWidth     =   5
               Height          =   2895
               Left            =   360
               Top             =   600
               Width           =   5295
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕÊÌ· »Ì‰ «·Œ“‰"
            Height          =   855
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   120
            Width           =   13335
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   7800
               TabIndex        =   99
               Top             =   255
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBoxTo 
               Height          =   315
               Left            =   150
               TabIndex        =   100
               Top             =   255
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·Œ“Ì‰Â «·„ÕÊ· „‰Â«"
               Height          =   345
               Index           =   3
               Left            =   11820
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·Œ“Ì‰Â «·„ÕÊ· «·ÌÂ«"
               Height          =   285
               Index           =   10
               Left            =   4740
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   285
               Width           =   1395
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1605
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   3480
            Width           =   14655
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   11760
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtNoteSerial2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   11760
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtNoteId2 
               Alignment       =   1  'Right Justify
               Height          =   495
               Left            =   12240
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   1080
               Visible         =   0   'False
               Width           =   735
            End
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   5130
               TabIndex        =   83
               Top             =   300
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   90
               TabIndex        =   84
               Top             =   270
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Index           =   7
               Left            =   10320
               TabIndex        =   85
               Top             =   240
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   661
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   -2147483637
            End
            Begin MSDataListLib.DataCombo DcboDebitSide1 
               Height          =   315
               Left            =   5160
               TabIndex        =   86
               Top             =   720
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide1 
               Height          =   315
               Left            =   90
               TabIndex        =   87
               Top             =   720
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Index           =   11
               Left            =   10320
               TabIndex        =   88
               Top             =   720
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   661
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   -2147483637
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   285
               Index           =   32
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   300
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   285
               Index           =   31
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   390
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ1:"
               Height          =   315
               Index           =   30
               Left            =   13170
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   330
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   285
               Index           =   29
               Left            =   5130
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   2220
               Width           =   975
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1050
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lblAccountInterval 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   1740
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰ 2 "
               Height          =   285
               Index           =   19
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   720
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰ 2 "
               Height          =   285
               Index           =   20
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   720
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ2:"
               Height          =   315
               Index           =   34
               Left            =   13170
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   810
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕÊÌ· »Ì‰ «·»‰Êﬂ"
            Height          =   1335
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   120
            Visible         =   0   'False
            Width           =   13335
            Begin VB.Frame Frame5 
               Caption         =   "„’«—Ì› ÕÊ«·… »‰ﬂÌ…"
               Height          =   615
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   600
               Width           =   5535
               Begin VB.TextBox txtTransferExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2400
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—”Ê„ «·ÕÊ«·Â"
                  Height          =   255
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin MSDataListLib.DataCombo DcBank 
               Height          =   315
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCBankTo 
               Height          =   315
               Left            =   7800
               TabIndex        =   76
               Top             =   240
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·»‰ﬂ «·„ÕÊ· „‰Â"
               Height          =   345
               Index           =   12
               Left            =   11820
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»‰ﬂ «·„ÕÊ· «·ÌÂ"
               Height          =   285
               Index           =   13
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   285
               Width           =   2235
            End
         End
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·‘Ìﬂ/«·ÕÊ«·Â"
            Height          =   1365
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1440
            Width           =   13365
            Begin VB.TextBox txtreport_no 
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
               Height          =   285
               Left            =   0
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Text            =   "1"
               Top             =   1320
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.TextBox txtperson 
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
               Height          =   285
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Tag             =   "·«»œ „‰ «œŒ«·  «·„” ›Ìœ"
               Top             =   270
               Width           =   4050
            End
            Begin VB.TextBox TxtChequeNumber 
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
               Height          =   285
               Left            =   10260
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Tag             =   "·«»œ „‰  —ﬁ„ «·‘Ìﬂ"
               Top             =   270
               Width           =   1545
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmBoxDrawing.frx":06D9
               Left            =   2280
               List            =   "FrmBoxDrawing.frx":06E9
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   2190
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox XPMTxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   0
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   60
               Top             =   720
               Width           =   11865
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   6600
               TabIndex        =   65
               Top             =   270
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               Format          =   238944257
               CurrentDate     =   39614
            End
            Begin VB.Label LblValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   405
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   720
               Width           =   8610
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„” ›Ìœ"
               Height          =   195
               Index           =   5
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   270
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
               Height          =   195
               Index           =   0
               Left            =   8220
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   270
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ//«·ÕÊ«·Â"
               Height          =   195
               Index           =   3
               Left            =   11805
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   270
               Width           =   1350
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰«¡ ⁄·Ï"
               Height          =   285
               Index           =   1
               Left            =   11880
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   720
               Width           =   1155
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   4800
         Left            =   15420
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   45
         Width           =   14685
         _cx             =   25903
         _cy             =   8467
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
         BackColor       =   -2147483633
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3750
            Left            =   120
            TabIndex        =   139
            Tag             =   "1"
            Top             =   120
            Width           =   14535
            _cx             =   25638
            _cy             =   6615
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmBoxDrawing.frx":0702
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
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   11040
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   5280
            Width           =   3375
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   4200
            Width           =   3375
         End
      End
   End
   Begin VB.TextBox TxtOderSerial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtOrder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox XPMTxtRemarks1 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   0
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Top             =   1080
      Width           =   4065
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11475
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TxtOrgValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   2130
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ComboBox CboDrawingType 
      Height          =   315
      ItemData        =   "FrmBoxDrawing.frx":0845
      Left            =   11460
      List            =   "FrmBoxDrawing.frx":0847
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1050
      Width           =   1815
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  ≈÷«›Ì…"
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
      Height          =   2595
      Index           =   0
      Left            =   9750
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   8820
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— »ﬂ‘› Õ”«» «·Œ“‰…"
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
         Height          =   1575
         Index           =   1
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   2145
         Begin MSComCtl2.DTPicker DtpBoxFrom 
            Height          =   330
            Left            =   90
            TabIndex        =   28
            Top             =   300
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   393216
            CalendarTrailingForeColor=   0
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   242024449
            CurrentDate     =   38845
         End
         Begin MSComCtl2.DTPicker DtpBoxTo 
            Height          =   360
            Left            =   90
            TabIndex        =   29
            Top             =   690
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   635
            _Version        =   393216
            CalendarTitleBackColor=   14737632
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   242024449
            CurrentDate     =   38845
         End
         Begin ImpulseButton.ISButton CmdShowReport 
            Cancel          =   -1  'True
            Height          =   405
            Left            =   90
            TabIndex        =   30
            Top             =   1080
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
            BackColor       =   14871017
            FontName        =   "Tahoma"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBoxDrawing.frx":0849
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Lab 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Lab 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   1740
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   720
            Width           =   315
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—’Ìœ «·Œ“‰… «·√‰"
         Height          =   315
         Index           =   8
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label LblBoxName 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   270
         Width           =   885
      End
      Begin VB.Label LblBoxAccount 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   690
         Width           =   975
      End
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   11490
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1410
      Width           =   1785
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10230
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5790
      Visible         =   0   'False
      Width           =   1275
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   14745
      _cx             =   26009
      _cy             =   1032
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " ÕÊÌ·«  „«·ÌÂ  "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Begin VB.TextBox oldtxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   1425
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBoxDrawing.frx":0BE3
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBoxDrawing.frx":0F7D
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBoxDrawing.frx":1317
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   3
         Left            =   615
         TabIndex        =   8
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBoxDrawing.frx":16B1
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   13890
      TabIndex        =   9
      Top             =   7140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   13020
      TabIndex        =   10
      Top             =   7140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   12270
      TabIndex        =   11
      Top             =   7140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   11535
      TabIndex        =   12
      Top             =   7140
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   10755
      TabIndex        =   13
      Top             =   7140
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   6
      Left            =   4140
      TabIndex        =   14
      Top             =   7140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   3210
      TabIndex        =   15
      Top             =   7140
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   3780
      TabIndex        =   23
      Top             =   7830
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   345
      Left            =   8700
      TabIndex        =   0
      Top             =   660
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   242024449
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   9930
      TabIndex        =   36
      Top             =   7140
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DcBranch 
      Height          =   315
      Left            =   0
      TabIndex        =   42
      Top             =   720
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   8
      Left            =   11880
      TabIndex        =   46
      Top             =   7140
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   5040
      TabIndex        =   49
      Top             =   7140
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«·„—›ﬁ« "
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   10
      Left            =   8760
      TabIndex        =   50
      Top             =   7140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·”‰œ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   9
      Left            =   7680
      TabIndex        =   51
      Top             =   7140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DcCostCenter 
      Bindings        =   "FrmBoxDrawing.frx":1A4B
      Height          =   315
      Left            =   5640
      TabIndex        =   52
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
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
   Begin ImpulseButton.ISButton Accredit 
      Height          =   375
      Left            =   240
      TabIndex        =   137
      Top             =   7140
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«—”«· ··«⁄ „«œ"
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorButton     =   -2147483635
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   6360
      TabIndex        =   142
      Top             =   7140
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "‰”Œ… „„«À·Â"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin XtremeSuiteControls.CheckBox IncludVAT 
      Height          =   270
      Left            =   1920
      TabIndex        =   143
      Top             =   1560
      Width           =   2175
      _Version        =   786432
      _ExtentX        =   3836
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   "«·ÕÊ«·…  ‘„· «·ﬁÌ„… «·„÷«›…"
      ForeColor       =   8388608
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰« ⁄·Ï ÿ·» ’—›"
      Height          =   285
      Index           =   37
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
      Height          =   255
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‘—Õ"
      Height          =   285
      Index           =   33
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   27
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   7680
      Width           =   7155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄ «·ﬁ«∆„ »«·⁄„·Ì…"
      Height          =   345
      Index           =   14
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·Œ“Ì‰Â «·„ÕÊ· „‰Â«"
      Height          =   345
      Index           =   11
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Lb_note_value_by_characters 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «· ÕÊÌ·"
      Height          =   285
      Index           =   9
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1050
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·⁄„·Ì…"
      Height          =   315
      Index           =   7
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   690
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   6
      Left            =   5790
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7830
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   285
      Index           =   5
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   810
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   7830
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   7830
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   345
      Index           =   0
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   765
      Width           =   1185
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7830
      Width           =   465
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1860
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7830
      Width           =   495
   End
End
Attribute VB_Name = "FrmBoxDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo(4)  As clsDCboSearch
Dim Line3 As Double
Dim Line2 As Double
Sub RetriveOrder(Optional OrderID As Double = 0, Optional OrderSerail As String)
Dim My_SQL As String
Dim rs2 As New ADODB.Recordset
If OrderID <> 0 Then
My_SQL = " select * from  TblExchange where id =" & OrderID & " "
Else
My_SQL = " select * from  TblExchange where  NoteSerial1='" & OrderSerail & "'  "
End If
My_SQL = My_SQL & " and type < 2 order by id"
 Set rs2 = New ADODB.Recordset
    rs2.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs2.RecordCount > 0 Then
    XPMTxtRemarks1.text = IIf(IsNull(rs2("Des").value), "", rs2("Des").value)
    TxtPerson2.text = IIf(IsNull(rs2("ToPerson").value), "", rs2("ToPerson").value)
    XPTxtVal.text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
    CboPayMentType.ListIndex = IIf(IsNull(rs2("Type").value), -1, rs2("Type").value)
    If val(CboPaymentType1.ListIndex) >= 2 Then
    CboPaymentType1.ListIndex = -1
    End If
   End If

End Sub
Public Function print_report(Optional NoteID As Double)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "SELECT    * from EXPENSES_ORDER2 "
 
    MySQL = MySQL & " Where ( NoteID = " & NoteID & ")   "
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "Expenses_order4.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "Expenses_order4.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text
        xReport.ParameterFields(7).AddCurrentValue DcboCreditSide.text
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text
        xReport.ParameterFields(7).AddCurrentValue DcboCreditSide.text
          
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
    SendTopost Me.Name, "Notes", "NoteID", 0, val(dcBranch.BoundText), val(XPTxtID.text), TxtNoteSerial1.text
  '' RsNetes.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
fillapprovData
End Sub

Private Sub CboDrawingType_Change()
 '  .AddItem "”Õ» ⁄«œÏ „‰ «·Œ“‰…"
 '           .AddItem " ÕÊÌ· „‰ Œ“‰… ≈·Ï ≈Œ—Ï"
 '           .AddItem " ÕÊÌ· „‰ »‰ﬂ ≈·Ï »‰ﬂ «Œ—"
 '           .AddItem " ÕÊÌ· „‰  ›—⁄ «·Ï ›—⁄ «Œ— "
      
    If Me.CboDrawingType.ListIndex = 0 Then '”Õ» ⁄«œÏ „‰ «·Œ“‰…
      If Me.TxtModFlg <> "R" Then txtTransferExpenses.text = 0
        Me.lbl(10).Visible = False
        Me.DcboBoxTo.Visible = False
        Frame1.Visible = True
        Frame2.Visible = False
        Frm2.Visible = False
        Frame3.Visible = False
    ElseIf Me.CboDrawingType.ListIndex = 1 Then ' ÕÊÌ· „‰ Œ“‰… ≈·Ï ≈Œ—Ï
    If Me.TxtModFlg <> "R" Then txtTransferExpenses.text = 0
        Me.lbl(10).Visible = True
        Me.DcboBoxTo.Visible = True
        Frame1.Visible = True
        Frame2.Visible = False
        Frm2.Visible = False
        Frame3.Visible = False
    ElseIf Me.CboDrawingType.ListIndex = 2 Then ' ÕÊÌ· „‰ »‰ﬂ ≈·Ï »‰ﬂ «Œ—
        Frame1.Visible = False
        Frame2.Visible = True
        Frm2.Visible = True
        Frame3.Visible = False
        
    ElseIf Me.CboDrawingType.ListIndex = 3 Then ' ÕÊÌ· „‰  ›—⁄ «·Ï ›—⁄ «Œ—
        Frame1.Visible = False
        Frame2.Visible = False
        Frm2.Visible = False
        Frame3.Visible = True
        'If Me.TxtModFlg <> "R" Then txtTransferExpenses.Text = 0
    End If

    If TxtModFlg.text <> "R" Then
        WriteDev
    End If

End Sub

Private Sub CboDrawingType_Click()
    CboDrawingType_Change
End Sub

Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.text = "E" Then
        Dcbank1.text = ""
        TxtChequeNumber1.text = ""
        Me.DcboBox1.text = ""
        DCAccounts1.text = ""

    End If
 
    If Me.CboPayMentType.ListIndex = 0 Then
        Me.lbl(21).Enabled = True
        Me.DcboBox1.Enabled = True
        Me.lbl(22).Enabled = False
        Me.lbl(23).Enabled = False
        Me.lbl(24).Enabled = False
        Me.lbl(17).Enabled = False
    
        Me.Dcbank1.Enabled = False
        Me.TxtChequeNumber1.Enabled = False
        Me.DtpChequeDueDate1.Enabled = False
 
        TxtChequeNumber1.text = ""
        Dcbank1.text = ""
 
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
    
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
        Me.lbl(22).Enabled = True
        Me.lbl(23).Enabled = True
        Me.lbl(24).Enabled = True
        Me.Dcbank1.Enabled = True
        Me.TxtChequeNumber1.Enabled = True
        Me.DtpChequeDueDate1.Enabled = True
        Me.lbl(17).Enabled = False
        Me.lbl(21).Enabled = False
    
        Me.DcboBox1.Enabled = False
        DcboBox1.text = ""
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
    
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
        Me.lbl(17).Enabled = True
        DCAccounts1.Enabled = True
 
        Me.lbl(21).Enabled = False
        Me.lbl(22).Enabled = False
        Me.lbl(23).Enabled = False
        Me.lbl(24).Enabled = False
        Me.Dcbank1.Enabled = False
        Me.TxtChequeNumber1.Enabled = False
        Me.DtpChequeDueDate1.Enabled = False
 
        TxtChequeNumber1.text = ""
    
        Me.DcboBox1.Enabled = False
        DcboBox1.text = ""
 
        Dcbank1.text = ""
    
    Else
        
        Me.lbl(17).Enabled = False
        Me.lbl(21).Enabled = False
        Me.lbl(22).Enabled = False
        Me.lbl(23).Enabled = False
        Me.lbl(24).Enabled = False
        Me.Dcbank1.Enabled = False
        Me.TxtChequeNumber1.Enabled = False
        Me.DtpChequeDueDate1.Enabled = False
 
        TxtChequeNumber1.text = ""
    
        Me.DcboBox1.Enabled = False
        DcboBox1.text = ""
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
    End If
If Me.TxtModFlg.text <> "R" Then
If val(CboPayMentType.ListIndex) <> -1 Then
CboPaymentType1.ListIndex = CboPayMentType.ListIndex
End If
End If
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboPaymentType1_Change()

    If Me.TxtModFlg.text = "E" Then
        Dcbank2.text = ""
   
        Me.DcboBox2.text = ""
        DCAccounts2.text = ""

    End If
 
    If Me.CboPaymentType1.ListIndex = 0 Then
        Me.lbl(36).Enabled = True
        Me.DcboBox2.Enabled = True
    
        Me.lbl(35).Enabled = False
        ' Me.lbl(15).Enabled = False
    
        Me.Dcbank2.Enabled = False
        Dcbank2.text = ""
 
        DCAccounts2.Enabled = False
        DCAccounts2.text = ""
    
    ElseIf Me.CboPaymentType1.ListIndex = 1 Then
        Me.lbl(35).Enabled = True
 
        Me.Dcbank2.Enabled = True
    
        Me.lbl(18).Enabled = False
        Me.lbl(36).Enabled = False
    
        Me.DcboBox2.Enabled = False
        DcboBox2.text = ""
        DCAccounts2.Enabled = False
        DCAccounts2.text = ""
    
    ElseIf Me.CboPaymentType1.ListIndex = 2 Then
        Me.lbl(18).Enabled = True
        DCAccounts2.Enabled = True
 
        Me.lbl(35).Enabled = False
        Me.lbl(36).Enabled = False
    
        Me.DcboBox2.Enabled = False
        DcboBox2.text = ""
 
        Dcbank2.text = ""
        Dcbank2.Enabled = False
    
    Else
        
        Me.lbl(18).Enabled = False
        Me.lbl(35).Enabled = False
        Me.lbl(36).Enabled = False
    
        Me.DcboBox2.Enabled = False
        DcboBox2.text = ""
        DCAccounts2.Enabled = False
        DCAccounts2.text = ""
        Dcbank2.text = ""
        Dcbank2.Enabled = False
   
    End If

End Sub

Private Sub CboPaymentType1_Click()
    CboPaymentType1_Change
End Sub

Private Sub Cmd_Click(index As Integer)
'   On Error GoTo ErrTrap

    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.rows = 1
            IncludVAT.value = vbUnchecked
            'XPTxtID.Text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=8"))
            Me.DCboUserName.BoundText = user_id
            Accredit.Caption = ""
            XPDtbTrans.SetFocus
            Me.dcBranch.BoundText = Current_branch
Me.dcBranch1.BoundText = Current_branch

        Case 1
        rs.Resync adAffectCurrent
        
                     If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
 
 
        If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
'rs.Resync adAffectCurrent
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            CuurentLogdata
CboDrawingType_Change
        Case 2
            If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
 
            my_branch = val(Me.dcBranch.BoundText)
If val(txtTransferExpensesBranch.text) + val(txtTransferExpenses.text) > 0 And IncludVAT.value = vbChecked Then
If GetValueAddedAccount(XPDtbTrans.value, , , 1, 23) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
            SaveData

        Case 3
       
            Undo

        Case 4
                     If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
 
 
        If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

           Load FrmBoxSearch
            FrmBoxSearch.SearchNoteType = 8
              FrmBoxSearch.show vbModal

        Case 6
            Unload Me

        Case 7
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
If SystemOptions.JLCodeBasedOnBranch = True Then
            ShowGL_cc Me.TxtNoteSerial.text, , 200, val(XPTxtID.text)
Else
    ShowGL_cc Me.TxtNoteSerial.text, , 200
End If

        Case 8
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DCBankTo.BoundText)), TxtNoteSerial.text

        Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber1.text, get_Cheque_report_no(val(Dcbank1.BoundText)), TxtNoteSerial.text

        Case 10

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text)
    
            End If
Case 11
      If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
If SystemOptions.JLCodeBasedOnBranch = True Then
ShowGL_cc Me.TxtNoteSerial2.text, , 200, val(Me.TxtNoteID2.text)
 
Else
ShowGL_cc Me.TxtNoteSerial2.text, , 200
  End If
            'print_Cheque TxtChequeNumber1.text, get_Cheque_report_no(Val(Dcbank1.BoundText)), TxtNoteSerial.text
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where ChqueNum='0'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '
    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then
    'GetMsgs 138, vbExclamation
    '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(XPMTxtRemarks.text)
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.text)
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.txtperson.text)
    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
    xReport.ParameterFields(15).AddCurrentValue Format$(DtpChequeDueDate.value, "dd/mm/yyyy")
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdAttach_Click()
     On Error Resume Next
           If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "0812201402"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub



Private Sub CmdShowReport_Click()
    Dim cBoxReport As ClsBoxesReports
    Dim Msg As String

    If Me.DcboBox.BoundText = "" Then
        Exit Sub
    Else
        Set cBoxReport = New ClsBoxesReports
        cBoxReport.BoxBalance Me.DcboBox.BoundText, Me.DtpBoxFrom.value, Me.DtpBoxTo.value
        Set cBoxReport = Nothing
    End If

End Sub

Private Sub DCAccounts1_Change()
    WriteDev
End Sub

Private Sub DCAccounts1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
      Account_search.case_id = 2602191
        Account_search.show
 
        
    End If

End Sub

Private Sub DcAccounts2_Change()
    WriteDev
End Sub

Private Sub DCAccounts2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
      Account_search.case_id = 260219
        Account_search.show
 
        
    End If


End Sub

Private Sub DcBank_Change()
    WriteDev
End Sub

Private Sub Dcbank1_Change()
    On Error Resume Next

    If Dcbank1.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & val(Dcbank1.BoundText)

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        If CboPayMentType.ListIndex = 1 Then
                     
            Me.DCAccounts1.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If
 
    End If
 
End Sub

Private Sub Dcbank1_Click(Area As Integer)
    Dcbank1_Change
End Sub

Private Sub Dcbank2_Change()
    On Error Resume Next

    If Dcbank2.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & val(Dcbank2.BoundText)

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        If CboPaymentType1.ListIndex = 1 Then
                     
            Me.DCAccounts2.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If
 
    End If

End Sub

Private Sub Dcbank2_Click(Area As Integer)
    Dcbank2_Change
End Sub

Private Sub DCBankTo_Change()
    WriteDev
End Sub

Private Sub DcboBox_Change()
    GetBoxData

    WriteDev
End Sub

Private Sub DcboBox1_Change()

    If DcboBox1.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DCAccounts1.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox1.BoundText))
    End If

End Sub

Private Sub DcboBox1_Click(Area As Integer)
    DcboBox1_Change
End Sub

Private Sub DcboBox2_Change()

    If DcboBox2.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DCAccounts2.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox2.BoundText))
    End If

End Sub

Private Sub DcboBox2_Click(Area As Integer)
    DcboBox2_Change
End Sub

Private Sub DcboBoxTo_Change()
    WriteDev
End Sub

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    CmdAttach.Caption = "Attachments"
    IncludVAT.Caption = "Include VAT"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Cmd(7).Caption = "Print JL "
    Cmd(11).Caption = "Print JL "
    Me.Caption = "Transfer Money"
    EleHeader.Caption = Me.Caption
Label2.Caption = "This Screen Allow Transfer Money with Many Methods"
Cmd(8).Caption = "Print Cheque"
    lbl(0).Caption = "OPR ID"
ISButton1.Caption = "Same Copy"
    lbl(7).Caption = "Date"
    lbl(8).Caption = "Type"
    lbl(5).Caption = "Value "
    lbl(3).Caption = "From Box"
lbl(33).Caption = "Remarks"
lbl(37).Caption = "Through Requested"
    
Frame5.Caption = "Cheque/Transfer Details"
Label3.Caption = "Value"

    lbl(10).Caption = "To Box"
    lbl(1).Caption = "Based ON"
    lbl(9).Caption = "OPR Type"
 Label8.Caption = "General Cent Cost"
    Fra(2).Caption = "GL"
    lbl(30).Caption = "GL#"
    lbl(34).Caption = "GL#"
    lbl(29).Caption = "Interval"

Frame3.Caption = "Transfer Between Branches"

lbl(16).Caption = "From Branch"
lbl(15).Caption = "To Branch"
lbl(25).Caption = "Payments Type"
lbl(28).Caption = "Recipet Type"
lbl(21).Caption = "From Box"
 lbl(36).Caption = "To Box"

lbl(22).Caption = "From Bank"
 lbl(35).Caption = "To Bank"
 
lbl(23).Caption = "Cheque No."
 lbl(24).Caption = "Due Date"
 
  lbl(17).Caption = "From ACC"
  
  lbl(18).Caption = "To ACC"
   lbl(26).Caption = "Recipet Name"
 Cmd(9).Caption = "Print Cheque"
 Cmd(10).Caption = "Print Vchr"
 Frame6.Caption = "Transfer Commisions Details"
 Label4.Caption = "Value"
 
 

    lbl(32).Caption = "Depit1"
    lbl(31).Caption = "Credit1"

  lbl(19).Caption = "Depit2"
    lbl(20).Caption = "Credit2"

Frm2.Caption = "Cheque\Transfer Details"
Label1(3).Caption = "No#"
Label1(0).Caption = "Due Date"
Label1(5).Caption = "Recipient Name"


    lbl(6).Caption = "By"
    lbl(2).Caption = "Curr. rec."
    lbl(4).Caption = "Rec. count."
    Fra(0).Caption = "Information"
    lbl(8).Caption = "Box Balance"

    Fra(1).Caption = "Box Report"
    Lab(4).Caption = "From"
    Lab(3).Caption = "To"

    CmdShowReport.Caption = "View"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    
    Frame1.Caption = "Transfer Between Boxes"
    Frame2.Caption = "Transfer Between banks"
    
    lbl(3).Caption = "From"
    lbl(10).Caption = "To"
        
    lbl(12).Caption = "From"
    lbl(13).Caption = "To"
    lbl(14).Caption = "Branch"
 
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    dcBranch1.BoundText = dcBranch.BoundText
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch1_Change()
LoadBranchData
If Me.TxtModFlg = "E" Or Me.TxtModFlg = "N" Then TxtNoteSerial.text = ""

    WriteDev
End Sub

Private Sub DcBranch1_Click(Area As Integer)
Dcbranch1_Change
End Sub

Private Sub DcBranch2_Change()
LoadBranchData2
If Me.TxtModFlg = "E" Or Me.TxtModFlg = "N" Then TxtNoteSerial2.text = ""

    WriteDev
End Sub

Private Sub DcBranch2_Click(Area As Integer)
DcBranch2_Change
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Dim StrSQL As String
    ScreenNameArabic = " ÕÊÌ·«  „«·ÌÂ"
    ScreenNameEnglish = "Bank Transfers"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 14
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboDrawingType
            .Clear
            .AddItem "Withdrawn From Box"
            .AddItem "Transfer between boxs"
            .AddItem "Transfer between Banks"
            .AddItem "Transfer between Branches"
     
        End With
        
                With Me.CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Cheque/Trans"
            .AddItem "Account"
        End With

        With Me.CboPaymentType1
            .Clear
        
            .AddItem "Cash"
            .AddItem "Cheque/Trans"
            .AddItem "Account"
        End With
        

    Else

        With Me.CboDrawingType
            .Clear
            .AddItem "”Õ» ⁄«œÏ „‰ «·Œ“‰…"
            .AddItem "  „‰ Œ“‰…/⁄Âœ… ≈·Ï ≈Œ—Ï"
            .AddItem "  „‰ »‰ﬂ ≈·Ï »‰ﬂ  "
            .AddItem "  „‰  ›—⁄ «·Ï ›—⁄  "
            
        End With

        With Me.CboPayMentType
            .Clear
            .AddItem "‰ﬁœÌ"
            .AddItem "‘Ìﬂ/ÕÊ«·Â"
            .AddItem "Õ”«»"
        End With

        With Me.CboPaymentType1
            .Clear
            .AddItem "‰ﬁœÌ"
            .AddItem "‘Ìﬂ /ÕÊ«·Â"
            .AddItem "Õ”«»"
        End With
        
    End If

    AddTip
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBoxes Me.DcboBox1
    Dcombos.GetBoxes Me.DcboBox2
Dcombos.GetCostCenter DcCostCenter
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboBox

    Dcombos.GetBoxes Me.DcboBoxTo
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboBoxTo

    Dcombos.GetBanks Me.Dcbank
    Dcombos.GetBanks Me.Dcbank1
    Dcombos.GetBanks Me.Dcbank2
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.Dcbank

    Dcombos.GetBanks Me.DCBankTo
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DCBankTo

    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBranches Me.dcBranch1
    

    Set cSearchDcbo(4) = New clsDCboSearch
    Set cSearchDcbo(4).Client = Me.dcBranch

    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetAccountingCodes Me.DcboDebitSide1
    Dcombos.GetAccountingCodes Me.DcboCreditSide1

    Dcombos.GetAccountingCodes Me.DCAccounts1, True
    Dcombos.GetAccountingCodes Me.DCAccounts2, True

    SetDtpickerDate Me.DtpBoxFrom
    SetDtpickerDate Me.DtpBoxTo
    SetDtpickerDate Me.XPDtbTrans

    Set cSearchDcbo(0) = New clsDCboSearch

    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where not( DrawingType is null  ) and NoteType=14"
    StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
    
    If SystemOptions.usertype <> UserAdminAll Then
 '       StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 14

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    Exit Sub
ErrTrap:
End Sub
Sub LoadBranchData()
Dim StrSQL As String

 Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo Dcbranch2
     Dcombos.ClearMyDataCombo Dcbank1
     Dcombos.ClearMyDataCombo DcboBox1
 
      If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT branch_id,branch_name From TblBranchesData"
    Else
        StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
    End If
 If Me.TxtModFlg <> "R" Then
    StrSQL = StrSQL & " Where branch_id <>" & val(Me.dcBranch1.BoundText) & " or branch_id=0 or branch_id is null"
 End If
     fill_combo Dcbranch2, StrSQL
     
     
   If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BankID,BankName From BanksData  "
    Else
        StrSQL = "Select BankID,BankNameE From BanksData   "
   End If
   If Me.TxtModFlg <> "R" Then
   StrSQL = StrSQL & " where BranchID =" & val(Me.dcBranch1.BoundText) & " or BranchID=0 or BranchID is null"
   End If
   fill_combo Me.Dcbank1, StrSQL
   
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BoxID,BoxName From tblBoxesData "
    Else
        StrSQL = "Select BoxID,BoxNamee From tblBoxesData "
   End If
   If Me.TxtModFlg <> "R" Then
    StrSQL = StrSQL & " where BranchId =" & val(Me.dcBranch1.BoundText) & ""
   End If
   fill_combo Me.DcboBox1, StrSQL
 End Sub
Sub LoadBranchData2()
Dim StrSQL As String

 Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo Dcbank2
     Dcombos.ClearMyDataCombo DcboBox2
   If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BankID,BankName From BanksData  "
    Else
        StrSQL = "Select BankID,BankNameE From BanksData   "
   End If
   StrSQL = StrSQL & " where BranchID =" & val(Me.Dcbranch2.BoundText) & " or BranchID=0 or BranchID is null"
   fill_combo Me.Dcbank2, StrSQL
   
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BoxID,BoxName From tblBoxesData "
    Else
        StrSQL = "Select BoxID,BoxNamee From tblBoxesData "
   End If
    StrSQL = StrSQL & " where BranchId =" & val(Me.Dcbranch2.BoundText) & " or BranchId=0 or BranchId is null"
   fill_combo Me.DcboBox2, StrSQL
 End Sub
 

Private Sub ISButton1_Click()
            TxtModFlg.text = "N"
            Me.XPTxtID.text = ""
 
            Me.DCboUserName.BoundText = user_id
              'Me.DcBranch.BoundText = Current_branch
     TxtNoteSerial.text = ""
     TxtNoteSerial1.text = ""
     


End Sub


Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "”Õ» „‰ «·Œ“‰…"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False

            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True

            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True

            XPTxtVal.locked = True
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True

            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False

            End If

        Case "N"
            '        Me.Caption = "”Õ» „‰ «·Œ“‰…( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False

            '     Me.XPBtnMove(0).Enabled = False
            '     Me.XPBtnMove(1).Enabled = False
            '     Me.XPBtnMove(2).Enabled = False
            '     Me.XPBtnMove(3).Enabled = False

            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
       
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "”Õ» „‰ «·Œ“‰…(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False

            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False

            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
        
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtOderSerial_Change()
If Me.TxtOderSerial.text <> "" Then
RetriveOrder , (Me.TxtOderSerial.text)
End If
End Sub

Private Sub TxtOderSerial_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF3 Then

                    
            TxtOderSerial.text = ""
            FrmReqExchangeSearch.show
            FrmReqExchangeSearch.lbltype = 4
        
      

    End If
End Sub

Private Sub TxtOrder_Change()
If Me.TxtOrder.text <> "" Then
RetriveOrder val(Me.TxtOrder.text)
End If
End Sub

Private Sub TxtPerson2_Change()
    txtperson.text = TxtPerson2.text
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteSerial1 As String)
'    On Error GoTo ErrTrap
    Dim RsTemp  As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

                    If rs.EOF Or rs.BOF Then
                        Exit Sub
                    End If
        End If
        
    If NoteSerial1 <> "" Then
          rs.Find "NoteSerial1=" & NoteSerial1, , adSearchForward, adBookmarkFirst

                    If rs.EOF Or rs.BOF Then
                        Exit Sub
                    End If
    End If
    
        
        
    End If
    
    
 If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
        Me.dcBranch1.BoundText = IIf(IsNull(rs("branch_no1").value), "", rs("branch_no1").value)
    Me.Dcbranch2.BoundText = IIf(IsNull(rs("branch_no2").value), "", rs("branch_no2").value)


    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
TxtNoteID2.text = IIf(IsNull(rs("NoteID2").value), "", (rs("NoteID2").value))
''//
Me.TxtOderSerial.text = IIf(IsNull(rs("OrderSerial").value), "", rs("OrderSerial").value)
Me.TxtOrder.text = IIf(IsNull(rs("OrderID").value), "", rs("OrderID").value)
''//
If Not IsNull(rs("IncludVAT").value) Then
If (rs("IncludVAT").value) = 1 Then
IncludVAT.value = vbChecked
Else
IncludVAT.value = vbUnchecked
End If
Else
IncludVAT.value = vbUnchecked
End If
Me.TxtNoteSerial2.text = IIf(IsNull(rs("NoteSerial2").value), "", rs("NoteSerial2").value)

    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
   If rs("ToCommision").value = vbTrue Then
        Me.ChkToCommision.value = vbChecked
    Else
        Me.ChkToCommision.value = vbUnchecked
    End If
    
    lbl(27).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    XPMTxtRemarks1.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    Me.Dcbank.BoundText = IIf(IsNull(rs("BankID").value), "", rs("BankID").value)
    txtTransferExpenses.text = IIf(IsNull(rs("TransferExpenses").value), "", (rs("TransferExpenses").value))
    txtTransferExpensesBranch.text = IIf(IsNull(rs("TransferExpensesBranch").value), "", (rs("TransferExpensesBranch").value))
    
    
    
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    '777777777777777777777777777777
    Me.Dcbank1.BoundText = IIf(IsNull(rs("BankID1").value), "", rs("BankID1").value)
    Me.Dcbank2.BoundText = IIf(IsNull(rs("BankID2").value), "", rs("BankID2").value)
    Me.DcboBox1.BoundText = IIf(IsNull(rs("BoxID1").value), "", rs("BoxID1").value)
    Me.DcboBox2.BoundText = IIf(IsNull(rs("BoxID2").value), "", rs("BoxID2").value)
    Me.TxtChequeNumber1.text = IIf(IsNull(rs("TxtChequeNumber1").value), "", rs("TxtChequeNumber1").value)
    Me.DtpChequeDueDate1.value = IIf(IsNull(rs("DtpChequeDueDate1").value), Date, rs("DtpChequeDueDate1").value)
    TxtPerson2.text = IIf(IsNull(rs("person").value), "", rs("person").value)

    '88888888888888888888888888888888888

    Me.DcboBoxTo.BoundText = IIf(IsNull(rs("BoxIDto").value), "", rs("BoxIDto").value)
    Me.DCBankTo.BoundText = IIf(IsNull(rs("BankIDto").value), "", rs("BankIDto").value)
    CboDrawingType.ListIndex = IIf(IsNull(rs("DrawingType").value), 0, rs("DrawingType").value)

    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), -1, rs("PaymentType").value)
    CboPaymentType1.ListIndex = IIf(IsNull(rs("PaymentType1").value), -1, rs("PaymentType1").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)

    Me.TxtChequeNumber.text = IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)
    Me.DtpChequeDueDate.value = IIf(IsNull(rs("DueDate").value), Date, rs("DueDate").value)
    txtperson.text = IIf(IsNull(rs("person").value), "", rs("person").value)

    Me.DCAccounts1.BoundText = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
    Me.DCAccounts2.BoundText = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
 
    Me.DcboDebitSide1.BoundText = getBranchCurrentAccount(val(Dcbranch2.BoundText))
    Me.DcboCreditSide1.BoundText = getBranchCurrentAccount(val(dcBranch1.BoundText))
     
    Me.DcboDebitSide.BoundText = DCAccounts2.BoundText
    Me.DcboCreditSide.BoundText = DCAccounts1.BoundText

    'Set RsTemp = New ADODB.Recordset
    'StrSQL = "Select * From NOTES Where NoteType=14 AND RetrunNoteID=" & Val(Me.XPTxtID.text)
    'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'If Not (RsTemp.BOF Or RsTemp.EOF) Then
    '    If Not IsNull(RsTemp("BoxID").value) Then
    '    Me.CboDrawingType.ListIndex = 1
    '    Me.DcboBoxTo.BoundText = RsTemp("BoxID").value
    '    Else
    '      Me.CboDrawingType.ListIndex = 2
    '    Me.DCBankTo.BoundText = RsTemp("BankiD").value
    '    End If
    '
    
    'Else
    '    Me.CboDrawingType.ListIndex = 0
    'End If

    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lblAccountInterval.Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If CboDrawingType.ListIndex < 3 Then
                    If RsDev("Credit_Or_Debit").value = 0 Then
                        Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                    ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                        Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                    End If
                End If

                RsDev.MoveNext
            Next i

        End If
    End If
fillapprovData
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtTransferExpenses_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtTransferExpenses.text, 0)
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim DblDif As Double
    Dim LngDevID As Long
    Dim RsDev As ADODB.Recordset
Dim Account_Code_dynamic   As String
Dim AccountVATCreit As String
Dim SngTemp3 As Double
Dim Percetage As Double
  Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
  '  On Error GoTo ErrTrap
       If (CboDrawingType.ListIndex = 2) And val(Me.txtTransferExpenses.text) > 0 Or ((CboDrawingType.ListIndex = 3) And val(Me.txtTransferExpensesBranch.text) > 0) Then
            Account_Code_dynamic = get_account_code_branch(52, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "No Branch Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»       „’—Ê›«   »‰ﬂÌ…  ›Ì  ‘«‘… —»ÿ «·Õ”«»«   ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "The bank Commisiion Account in this Branch is not specific", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If
        
        End If
        
    If Me.TxtModFlg.text <> "R" Then
        If Me.CboDrawingType.ListIndex = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·⁄„·Ì…  ...!!! "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboDrawingType.SetFocus
            Exit Sub
        End If
    
        If CboDrawingType.ListIndex = 0 Or CboDrawingType.ListIndex = 1 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

        ElseIf CboDrawingType.ListIndex = 2 Then

            If Trim(Me.Dcbank.BoundText) = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ «·„ÕÊ· „‰Â..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbank.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
        If val(XPTxtVal.text) = 0 Then
            Msg = "ÌÃ» «œŒ«· «·ﬁÌ„Â       "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            Exit Sub
        End If
    
        If Me.CboDrawingType.ListIndex = 1 Then
            If val(Me.DcboBoxTo.BoundText) = 0 Then
                Msg = "ÌÃ»  ÕœÌœ «·Œ“‰… «·„ÕÊ· ·Â«...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If Me.DcboBoxTo.Visible = True Then
                    DcboBoxTo.SetFocus
                End If

                Exit Sub
            ElseIf val(Me.DcboBox.BoundText) = val(Me.DcboBoxTo.BoundText) Then
                Msg = "·«Ì„ﬂ‰ «· ÕÊÌ· »Ì‰ ‰›” «·Œ“‰‹‹‹… ....!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If
    
        If Me.CboDrawingType.ListIndex = 2 Then
            If val(Me.DCBankTo.BoundText) = 0 Then
                Msg = "ÌÃ»  ÕœÌœ «·»‰ﬂ «·„ÕÊ· ·Â...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If Me.DCBankTo.Visible = True Then
                    DCBankTo.SetFocus
                End If

                Exit Sub
            ElseIf val(Me.Dcbank.BoundText) = val(Me.DCBankTo.BoundText) Then
                Msg = "·«Ì„ﬂ‰ «· ÕÊÌ· »Ì‰ ‰›” «·»‰ﬂ ....!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
            
            
            
                If Trim$(Me.TxtChequeNumber.text) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ./«·ÕÊ«·Â..!!"
                    Else
                        Msg = "Enter Cheque No:...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If


            
        End If
    
        If Me.CboDrawingType.ListIndex = 3 Then
                        
            If val(Me.dcBranch1.BoundText) = 0 Then
                Msg = "ÌÃ»  ÕœÌœ «·›—⁄  «·„ÕÊ· „‰…...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If Me.dcBranch1.Visible = True Then
                    dcBranch1.SetFocus
                    Sendkeys "{F4}"
                End If

                Exit Sub
                            
            End If
                        
            If val(Me.Dcbranch2.BoundText) = 0 Then
                Msg = "ÌÃ»  ÕœÌœ «·›—⁄  «·„ÕÊ· «·Ì… ...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If Me.Dcbranch2.Visible = True Then
                    Dcbranch2.SetFocus
                    Sendkeys "{F4}"
                End If

                Exit Sub
                            
            End If
                    
            If Me.CboPayMentType.ListIndex = -1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
                Else
                    Msg = "Select Payment method ...!!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboPayMentType.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
                    
            If Me.CboPayMentType.ListIndex = 0 Then
                If Trim(Me.DcboBox1.BoundText) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…   «·„ÕÊ·  „‰Â« ..!!"
                    Else
                        Msg = "Select Box..!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboBox1.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If

            ElseIf Me.CboPayMentType.ListIndex = 1 Then

                If Me.Dcbank1.BoundText = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ «·„ÕÊ· „‰Â...!!"
                    Else
                        Msg = "Select Bank...!!"
                    
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Dcbank1.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
        
                If Trim$(Me.TxtChequeNumber1.text) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    Else
                        Msg = "Enter Cheque No:...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber1.SetFocus
                    Exit Sub
                End If
         
            ElseIf Me.CboPayMentType.ListIndex = 2 Then

                If (Me.DCAccounts1.BoundText) = "" Then
                    Msg = "ÌÃ»  ÕœÌœ «·Õ”«»  «·„ÕÊ· „‰…...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                    If Me.DCAccounts1.Visible = True Then
                        DCAccounts1.SetFocus
                        Sendkeys "{F4}"
                    End If

                    Exit Sub
                            
                End If
                        
            End If

            'part 222222222222
            If Me.CboPaymentType1.ListIndex = -1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«—   «·«” ·«„  ...!!!"
                Else
                    Msg = "Select Payment method ...!!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboPaymentType1.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
                    
            If Me.CboPaymentType1.ListIndex = 0 Then
                If Trim(Me.DcboBox2.BoundText) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…   «·„ÕÊ·  «·ÌÂ« ..!!"
                    Else
                        Msg = "Select Box..!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboBox2.SetFocus
                     Sendkeys "{F4}"
                    Exit Sub
                End If

            ElseIf Me.CboPaymentType1.ListIndex = 1 Then

                If Me.Dcbank2.BoundText = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ «·„ÕÊ· «·ÌÂ«...!!"
                    Else
                        Msg = "Select Bank...!!"
                    
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Dcbank2.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
         
            ElseIf Me.CboPaymentType1.ListIndex = 2 Then

                If (Me.DCAccounts2.BoundText) = "" Then
                    Msg = "ÌÃ»  ÕœÌœ «·Õ”«»  «·„ÕÊ· «·ÌÂ«...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                    If Me.DCAccounts2.Visible = True Then
                        DCAccounts2.SetFocus
                        Sendkeys "{F4}"
                    End If

                    Exit Sub
                            
                End If
                        
            End If
                        
            If (Me.DCAccounts2.BoundText) = "" Then
                Msg = "ÌÃ»  ÕœÌœ «·Õ”«»  «·„ÕÊ· «·Ì… ...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If Me.DCAccounts2.Visible = True Then
                    DCAccounts2.SetFocus
                End If

                Exit Sub
                            
            End If
                    
        End If
    
        If Me.TxtModFlg.text = "N" Then
            '«· «ﬂœ „‰ «‰ —’Ìœ «·Œ“‰… Ì”„Õ »Œ—ÊÃ Â–« «·„»·€
            '   If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), XPDtbTrans.value, True) = False Then
            '       Exit Sub
            '   End If
        ElseIf Me.TxtModFlg.text = "E" Then
            '«· «ﬂœ „‰ «‰ —’Ìœ «·Œ“‰… Ì”„Õ »Œ—ÊÃ Â–« «·„»·€
         End If
    my_branch = val(dcBranch.BoundText)
        If TxtNoteSerial.text = "" Then
        
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                    If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                    End If
            End If
            
        End If
       '*************************
'If val(CboDrawingType.ListIndex) = 3 Then
'              If SystemOptions.JLCodeBasedOnBranch = True Then
            
            
'             End If

'End If

        
        If TxtNoteSerial1.text = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 16, 14) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ œ›⁄ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbTrans.value, 16, 14) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 16, 14)
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True
    
        If TxtModFlg.text = "N" Then
            Me.XPTxtID.text = new_id("NOTES", "NoteID", "")
            rs.AddNew
            rs("NoteID").value = val(XPTxtID.text)
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
         
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID2.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            If val(TxtNoteID2.text) <> val(Me.XPTxtID.text) Then
         StrSQL = "Delete From notes  Where  NoteType=14 and NoteID=" & val(TxtNoteID2.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            End If
           
        End If
        
        
          If val(CboDrawingType.ListIndex) = 3 Then
          Dim resultstr As String
          Dim des As String
          
               des = "  »‰«¡ ⁄·Ì ”‰œ  ÕÊÌ·  »—ﬁ„  " & TxtNoteSerial1 & " „‰ «·›—⁄  " & dcBranch1.text
               If TxtNoteSerial2.text = "" Then
                  resultstr = Notes_coding(val(Dcbranch2.BoundText), XPDtbTrans.value)
            If resultstr = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«   ··›—⁄ «·„ÕÊ· «·Ì…": Exit Sub
            Else
                       
                If resultstr = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ ··›—⁄ «·„ÕÊ· «·Ì… ": Exit Sub
                Else
                    TxtNoteSerial2.text = resultstr
                End If
            End If
        End If
      rs.update
        Dim NoteID As Long
        If SystemOptions.JLCodeBasedOnBranch = True Then
        CreateNotes NoteID, (XPDtbTrans.value), val(Dcbranch2.BoundText), 14, val(XPTxtVal.text), TxtNoteSerial2, TxtNoteSerial1, , , , des, ToHijriDate(XPDtbTrans.value)
        Else
        'TxtNoteSerial2 = TxtNoteSerial
       CreateNotes NoteID, (XPDtbTrans.value), val(dcBranch.BoundText), 14, val(XPTxtVal.text), TxtNoteSerial, TxtNoteSerial1, , , , des, ToHijriDate(XPDtbTrans.value)
          End If
           Me.TxtNoteID2.text = NoteID
          End If
          
         If val(CboDrawingType.ListIndex) = 3 Then
         rs("NoteID2").value = IIf(Me.TxtNoteID2.text = "", Null, val(TxtNoteID2.text))
         rs("NoteSerial2").value = IIf(Me.TxtNoteSerial2.text = "", Null, Me.TxtNoteSerial2.text)
         Else
          rs("NoteID2").value = Null
         rs("NoteSerial2").value = Null
         End If
         '''/
            rs("OrderSerial").value = IIf(Me.TxtOderSerial.text = "", Null, Me.TxtOderSerial.text)
    rs("OrderID").value = IIf(Me.TxtOrder.text = "", Null, Me.TxtOrder.text)
         ''//
       rs("TransferExpenses").value = val(txtTransferExpenses.text)
       rs("TransferExpensesBranch").value = val(txtTransferExpensesBranch.text)
       
       'ChkToCommision
          If Me.ChkToCommision.value = vbChecked Then
        rs("ToCommision").value = 1
    ElseIf Me.ChkToCommision.value = vbUnchecked Then
        rs("ToCommision").value = 0
    End If
        If IncludVAT.value = vbChecked Then
        rs("IncludVAT").value = 1
        Else
        rs("IncludVAT").value = 0
        End If
    
        rs("DrawingType").value = IIf(Me.CboDrawingType.ListIndex = -1, Null, CboDrawingType.ListIndex)
    
        rs("PaymentType").value = IIf(Me.CboPayMentType.ListIndex = -1, Null, CboPayMentType.ListIndex)
        rs("PaymentType1").value = IIf(Me.CboPaymentType1.ListIndex = -1, Null, CboPaymentType1.ListIndex)

        rs("NoteSerial").value = IIf(Me.TxtNoteSerial.text = "", Null, Me.TxtNoteSerial.text)
    
        rs("Note_Value").value = val(Me.XPTxtVal.text)
        Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.XPTxtVal.text + val(Me.txtTransferExpenses.text), "0.00"), 0, True, ".")
        rs("note_value_by_characters").value = Trim$(Me.Lb_note_value_by_characters.Caption)

        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
             'rs("Remark").value = "    ÕÊÌ·«  „«·ÌÂ " & " »—ﬁ„  " & Me.TxtNoteSerial1.text & "   " & IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
             
        rs("CusID").value = Null
     rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("DueDate").value = Null
        rs("Transaction_ID").value = Null
        rs("Member_ID").value = Null
        rs("ExpensesID").value = Null
    
        rs("NoteType").value = 14
        rs("NoteDate").value = XPDtbTrans.value
        rs("UserID").value = user_id

        If CboDrawingType.ListIndex = 0 Then
            rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
            rs("BoxIDTo").value = Null
            rs("BankID").value = Null
            rs("BankIDTo").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("person").value = Null
            rs("branch_no1").value = Null
            rs("branch_no2").value = Null
            rs("Account_Code1").value = Null
            rs("Account_Code2").value = Null
    
        ElseIf CboDrawingType.ListIndex = 1 Then
            rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
            rs("BoxIDTo").value = IIf(Me.DcboBoxTo.BoundText = "", Null, Me.DcboBoxTo.BoundText)
            rs("BankID").value = Null
            rs("BankIDTo").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("person").value = Null
      
            rs("branch_no1").value = Null
            rs("branch_no2").value = Null
            rs("Account_Code1").value = Null
            rs("Account_Code2").value = Null
        ElseIf CboDrawingType.ListIndex = 2 Then
            rs("BankID").value = IIf(Me.Dcbank.BoundText = "", Null, Me.Dcbank.BoundText)
            rs("BankIDTo").value = IIf(Me.DCBankTo.BoundText = "", Null, Me.DCBankTo.BoundText)
            rs("BoxID").value = Null
            rs("BoxIDto").value = Null
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("person").value = IIf(Me.txtperson.text = "", "", Me.txtperson.text)
        
            rs("branch_no1").value = Null
            rs("branch_no2").value = Null
            rs("Account_Code1").value = Null
            rs("Account_Code2").value = Null
     
        ElseIf CboDrawingType.ListIndex = 3 Then
            rs("BoxID").value = Null
            rs("BoxIDto").value = Null
            rs("BankID").value = Null
            rs("BankIDTo").value = Null
        
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("person").value = IIf(Me.txtperson.text = "", "", Me.txtperson.text)
      
            rs("branch_no1").value = IIf(Me.dcBranch1.BoundText = "", Null, Me.dcBranch1.BoundText)
            rs("branch_no2").value = IIf(Me.Dcbranch2.BoundText = "", Null, Me.Dcbranch2.BoundText)
            rs("Account_Code1").value = IIf(Me.DCAccounts1.BoundText = "", Null, Me.DCAccounts1.BoundText)
            rs("Account_Code2").value = IIf(Me.DCAccounts2.BoundText = "", Null, Me.DCAccounts2.BoundText)
    
            If CboPayMentType.ListIndex = 0 Then
                rs("BoxID1").value = IIf(Me.DcboBox1.BoundText = "", Null, Me.DcboBox1.BoundText)
                rs("BankID1").value = Null
  
                rs("TxtChequeNumber1").value = Null
                rs("DueDate").value = Null
                rs("person").value = Null
      
      '****************************
                rs("TxtChequeNumber1").value = Trim$(Me.TxtChequeNumber1.text)
                rs("DueDate").value = Me.DtpChequeDueDate.value
                rs("person").value = IIf(Me.TxtPerson2.text = "", "", Me.TxtPerson2.text)
     '****************************
     
            ElseIf CboPayMentType.ListIndex = 1 Then
                rs("BankID1").value = IIf(Me.Dcbank1.BoundText = "", Null, Me.Dcbank1.BoundText)
 
                rs("TxtChequeNumber1").value = Trim$(Me.TxtChequeNumber1.text)
                rs("DueDate").value = Me.DtpChequeDueDate.value
                rs("person").value = IIf(Me.TxtPerson2.text = "", "", Me.TxtPerson2.text)
                rs("BoxID").value = Null
              '****************************
                rs("TxtChequeNumber1").value = Trim$(Me.TxtChequeNumber1.text)
                rs("DueDate").value = Me.DtpChequeDueDate.value
                rs("person").value = IIf(Me.TxtPerson2.text = "", "", Me.TxtPerson2.text)
     '****************************
     
            ElseIf CboPayMentType.ListIndex = 2 Then
                rs("BoxID1").value = Null
                rs("BankID1").value = Null
                rs("TxtChequeNumber1").value = Null
                rs("DueDate").value = Null
                rs("person").value = Null
          '****************************
                rs("TxtChequeNumber1").value = Trim$(Me.TxtChequeNumber1.text)
                rs("DueDate").value = Me.DtpChequeDueDate.value
                rs("person").value = IIf(Me.TxtPerson2.text = "", "", Me.TxtPerson2.text)
     '****************************
       
       
            End If
    
            '«·Ã“¡ «·„ÕÊ· «·Ì…
            If CboPaymentType1.ListIndex = 0 Then
                rs("BoxID2").value = IIf(Me.DcboBox2.BoundText = "", Null, Me.DcboBox2.BoundText)
                rs("BankID2").value = Null
      
            ElseIf CboPaymentType1.ListIndex = 1 Then
                rs("BankID2").value = IIf(Me.Dcbank2.BoundText = "", Null, Me.Dcbank2.BoundText)
                rs("BoxID2").value = Null
        
            ElseIf CboPaymentType1.ListIndex = 2 Then
                rs("BoxID2").value = Null
                rs("BankID2").value = Null
      
            End If
    
        End If

        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(16) ' ÕÊÌ·«  „«·ÌÂ
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
     
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
 
        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
        rs("note_value_by_characters").value = IIf(Lb_note_value_by_characters = "", Null, Lb_note_value_by_characters)
    
        rs.update
    
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
       '     RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                  StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
           
            '«·ÿ—› «·„œÌ‰
            Dim StrDes As String
            Dim lineno As Integer
             lineno = 1
            StrDes = "    ÕÊÌ·«  „«·ÌÂ " & " »—ﬁ„  " & Me.TxtNoteSerial1.text & "   "
            
            RsDev.AddNew
        
            If CboDrawingType.ListIndex = 0 Or CboDrawingType.ListIndex = 1 Then
                RsDev("branch_id").value = val(Me.dcBranch.BoundText) '  GeBranchInfo("TblBoxesData", "boxid", val(Me.DcboBox.BoundText))
        
            ElseIf CboDrawingType.ListIndex = 2 Then
                RsDev("branch_id").value = val(Me.dcBranch.BoundText) ' GeBranchInfo("BanksData", "bankid", val(Me.DcBank.BoundText))
        
            ElseIf CboDrawingType.ListIndex = 3 Then
                RsDev("branch_id").value = val(Dcbranch2.BoundText)
            End If
        
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            'RsDev("DEV_ID_Line_No").value = LineNo1
                Line2 = setfoxy_Line
                  RsDev("DEV_ID_Line_No").value = lineno
                  If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
                  
             RsDev("DEV_ID_Line_No1").value = Line2
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text
            
             If CboDrawingType.ListIndex = 3 Then
            RsDev("Notes_ID").value = val(Me.TxtNoteID2.text)
            
            Else
             RsDev("Notes_ID").value = val(XPTxtID.text)
            End If
            
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = val(Me.DCboUserName.BoundText)
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
        
            If CboDrawingType.ListIndex = 3 Then
                lineno = lineno + 1
                StrDes = "    ÕÊÌ·«  „«·ÌÂ " & " »—ﬁ„  " & Me.TxtNoteSerial1.text & "   "
                RsDev.AddNew
  
                RsDev("branch_id").value = val(dcBranch1.BoundText)
                  If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
                'RsDev("Posted").value = Posted
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                  Line3 = setfoxy_Line
                 RsDev("DEV_ID_Line_No1").value = Line3

                RsDev("Account_Code").value = Me.DcboDebitSide1.BoundText
                RsDev("NextAccount_Code").value = Me.DcboCreditSide1.BoundText
                RsDev("Value").value = val(Me.XPTxtVal.text)
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev.update
       
            End If



            lineno = lineno + 1
            
 '«·ÿ—› «·„œÌ‰
            ' ›Ì Õ«·… «·ÕÊ«·«  «·»‰ﬂÌ… ÊÊÃÊœ „’—Ê›«  »‰ﬂÌ… ⁄·Ì⁄«
            If (CboDrawingType.ListIndex = 2) And val(Me.txtTransferExpenses.text) > 0 Then
                RsDev.AddNew
                If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
              '  RsDev("Posted").value = Posted
                       If IncludVAT.value = vbChecked Then
                          GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                        If Percetage = 0 Then
                           Percetage = 1
                        End If
                         Percetage = Percetage / 100 + 1
                         SngTemp3 = val(Me.txtTransferExpenses.text) / Percetage
                      Else
                         SngTemp3 = val(Me.txtTransferExpenses.text)
                     End If
                     
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                RsDev("DEV_ID_Line_No1").value = lineno
                RsDev("Account_Code").value = Account_Code_dynamic
                RsDev("Value").value = SngTemp3
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            
                RsDev.update
               If IncludVAT.value = vbChecked Then
                lineno = lineno + 1
        ''/////////////////
             RsDev.AddNew
                If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
              '  RsDev("Posted").value = Posted
             
                GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                SngTemp3 = SngTemp3 * Percetage / 100
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                RsDev("DEV_ID_Line_No1").value = lineno
                RsDev("Account_Code").value = AccountVATCreit
                RsDev("Value").value = SngTemp3
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & "Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            
                RsDev.update
            End If
            End If
            lineno = lineno + 1
            '«·ÿ—› «·œ«∆‰
            RsDev.AddNew
                  If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
'RsDev("Posted").value = Posted
            If CboDrawingType.ListIndex = 0 Then
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            ElseIf CboDrawingType.ListIndex = 1 Then
                RsDev("branch_id").value = val(Me.dcBranch.BoundText) ' GeBranchInfo("TblBoxesData", "boxid", val(Me.DcboBoxTo.BoundText))
            ElseIf CboDrawingType.ListIndex = 2 Then
                RsDev("branch_id").value = val(Me.dcBranch.BoundText) 'GeBranchInfo("BanksData", "bankid", val(Me.DCBankTo.BoundText))
        
            ElseIf CboDrawingType.ListIndex = 3 Then
                RsDev("branch_id").value = val(Me.dcBranch1.BoundText)
         
            End If
        
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("NextAccount_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text) + val(txtTransferExpenses.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
        
            
            If CboDrawingType.ListIndex = 3 Then
                lineno = lineno + 1
                RsDev.AddNew
                  If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
           ' RsDev("Posted").value = Posted
                RsDev("branch_id").value = val(Dcbranch2.BoundText)
        
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                RsDev("Account_Code").value = Me.DcboCreditSide1.BoundText
                RsDev("NextAccount_Code").value = Me.DcboDebitSide1.BoundText
                RsDev("Value").value = val(Me.XPTxtVal.text) + val(txtTransferExpenses.text)
                RsDev("Credit_Or_Debit").value = 1
                RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text
                RsDev("Notes_ID").value = val(TxtNoteID2.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev.update
            End If
            
            
                        lineno = lineno + 1
            
 '«·ÿ—› «·„œÌ‰
            '  ··›—Ê⁄ ›Ì Õ«·… «·ÕÊ«·«  «·»‰ﬂÌ… ÊÊÃÊœ „’—Ê›«  »‰ﬂÌ… ⁄·Ì⁄«
            If (CboDrawingType.ListIndex = 3) And val(Me.txtTransferExpensesBranch.text) > 0 Then
              
                   
                
            If ChkToCommision.value = vbUnchecked Then ' ⁄·Ì «·„ÕÊ· „‰…
                        RsDev.AddNew
                 If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
                  If IncludVAT.value = vbChecked Then
                          GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                        If Percetage = 0 Then
                           Percetage = 1
                        End If
                         Percetage = Percetage / 100 + 1
                         SngTemp3 = val(Me.txtTransferExpensesBranch.text) / Percetage
                      Else
                         SngTemp3 = val(Me.txtTransferExpensesBranch.text)
                     End If
                    '    RsDev("Posted").value = Posted
                        RsDev("branch_id").value = val(Me.dcBranch1.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
              
                        RsDev("Account_Code").value = Account_Code_dynamic
                        RsDev("Value").value = SngTemp3
                        RsDev("Credit_Or_Debit").value = 0
                        RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text & " „’«—Ì› ÕÊ«·… »‰ﬂÌ… "
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    
                        RsDev.update
              If IncludVAT.value = vbChecked Then
                                lineno = lineno + 1
                    ''////////////////////////////////
                    GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                SngTemp3 = SngTemp3 * Percetage / 100
                                   RsDev.AddNew
                 If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
                    '    RsDev("Posted").value = Posted
                        RsDev("branch_id").value = val(Me.dcBranch1.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
              
                        RsDev("Account_Code").value = AccountVATCreit
                        RsDev("Value").value = SngTemp3
                        RsDev("Credit_Or_Debit").value = 0
                        RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text & "Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    
                        RsDev.update
                   End If
                                lineno = lineno + 1
                                
                       RsDev.AddNew
                  If Posted = 1 Then
                  RsDev("Posted").value = 1
                  Else
                  RsDev("Posted").value = Null
                  End If
                      ' RsDev("Posted").value = Posted
                        RsDev("branch_id").value = val(Me.dcBranch1.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
                        RsDev("Account_Code").value = Me.DCAccounts1.BoundText
                        RsDev("Value").value = val(Me.txtTransferExpensesBranch.text)
                        RsDev("Credit_Or_Debit").value = 1
                        RsDev("Double_Entry_Vouchers_Description").value = StrDes & CHR(13) & XPMTxtRemarks.text & " „’«—Ì› ÕÊ«·… »‰ﬂÌ… "
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    
                        RsDev.update
                        
   Else '"⁄·Ì «·„ÕÊ· „‰…
       GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
               If IncludVAT.value = vbChecked Then
                
                           GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                        If Percetage = 0 Then
                           Percetage = 1
                        End If
                         Percetage = Percetage / 100 + 1
                         SngTemp3 = val(Me.txtTransferExpensesBranch.text) / Percetage
                      Else
                         SngTemp3 = val(Me.txtTransferExpensesBranch.text)
                     End If
                
            RsDev.AddNew
                        RsDev("branch_id").value = val(Me.Dcbranch2.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
                        RsDev("Account_Code").value = Account_Code_dynamic
                        RsDev("Value").value = SngTemp3
                        RsDev("Credit_Or_Debit").value = 0
                        RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & "Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                  If Posted = 1 Then
                        RsDev("Posted").value = 1
                  Else
                        RsDev("Posted").value = Null
                  End If
                  '  RsDev("Posted").value = Posted
                        RsDev.update
                        
                                
                     If IncludVAT.value = vbChecked Then
                     lineno = lineno + 1
                        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                       SngTemp3 = SngTemp3 * Percetage / 100
              
                        RsDev.AddNew
                        RsDev("branch_id").value = val(Me.Dcbranch2.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
                        RsDev("Account_Code").value = AccountVATCreit
                        RsDev("Value").value = SngTemp3
                        RsDev("Credit_Or_Debit").value = 0
                        RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                  If Posted = 1 Then
                        RsDev("Posted").value = 1
                  Else
                        RsDev("Posted").value = Null
                  End If
                  '  RsDev("Posted").value = Posted
                        RsDev.update
                 End If
                        
                                lineno = lineno + 1
                        
                       RsDev.AddNew
                        RsDev("branch_id").value = val(Me.dcBranch1.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
                        RsDev("Account_Code").value = Me.DCAccounts1.BoundText
                        RsDev("Value").value = val(Me.txtTransferExpensesBranch.text)
                        RsDev("Credit_Or_Debit").value = 1
                        RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                     If Posted = 1 Then
                        RsDev("Posted").value = 1
                     Else
                         RsDev("Posted").value = Null
                     End If
                  '  RsDev("Posted").value = Posted
                        RsDev.update
                        
                           lineno = lineno + 1
                            RsDev.AddNew
                        RsDev("branch_id").value = val(Me.dcBranch1.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
                        RsDev("Account_Code").value = DcboDebitSide1.BoundText
                        RsDev("NextAccount_Code").value = DcboCreditSide1.BoundText
                        RsDev("Value").value = val(Me.txtTransferExpensesBranch.text)
                        RsDev("Credit_Or_Debit").value = 0
                        RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                   If Posted = 1 Then
                        RsDev("Posted").value = 1
                  Else
                        RsDev("Posted").value = Null
                  End If
                    'RsDev("Posted").value = Posted
                        RsDev.update
                        
                                lineno = lineno + 1
                        
                       RsDev.AddNew
                        RsDev("branch_id").value = val(Me.Dcbranch2.BoundText)
                        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                        RsDev("DEV_ID_Line_No").value = lineno
                        RsDev("DEV_ID_Line_No1").value = lineno
                        RsDev("Account_Code").value = Me.DcboCreditSide1.BoundText
                        RsDev("NextAccount_Code").value = Me.DcboDebitSide1.BoundText
                        RsDev("Value").value = val(Me.txtTransferExpensesBranch.text)
                        RsDev("Credit_Or_Debit").value = 1
                        RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
                        RsDev("Notes_ID").value = val(XPTxtID.text)
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                  If Posted = 1 Then
                        RsDev("Posted").value = 1
                  Else
                        RsDev("Posted").value = Null
                  End If
                  '  RsDev("Posted").value = Posted
                        RsDev.update
   
   
   
   
                        
                End If
            End If
            
            
        
            LblDevID.Caption = LngDevID
            lblAccountInterval.Caption = SystemOptions.SysCurrentAccountIntervalID
        End If

        Cn.CommitTrans
        BeginTrans = False
        GetBoxData
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
                lbl(27).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
        
        End Select

        TxtModFlg.text = "R"
        Retrive
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, " ÕÊÌ·«  „«·ÌÂ", Me.XPDtbTrans.value, DcboDebitSide.BoundText, DcboDebitSide.text
       
     '  If val(CboDrawingType.ListIndex) = 3 Then
     '     save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, " ÕÊÌ·«  „«·ÌÂ", Me.XPDtbTrans.value, DcboDebitSide1.BoundText, DcboDebitSide1.text
     '     End If
fillapprovData
 updateNotesValueAndNobytext val(TxtNoteID2.text)
 updateNotesValueAndNobytext val(XPTxtID.text)
    End If

    Exit Sub
ErrTrap:
    
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    
    If Not (rs.BOF Or rs.EOF) Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If
    End If
    
    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub


Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date, DebitSideID As String, _
                                         DebitSide As String) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND (  kedno =" & val(XPTxtID.text) & " or  kedno =" & val(Me.TxtNoteID2.text) & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
        
    'rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'ÿ—› „œÌ‰
    rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = val(XPTxtVal.text)
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = val(XPTxtID.text)
    rs("kedno").value = val(XPTxtID.text)
        
    rs("opr_type").value = opr_type
    rs("account_name").value = DebitSide
    rs("account_no").value = DebitSideID
  '  rs("line_no").value = 1
     rs("line_no").value = Line2
     
          rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtNoteSerial.text
        rs("Remark").value = " ÕÊÌ·«  „«·Ì… »—ﬁ„ " & TxtNoteSerial1 & "    " & Me.XPMTxtRemarks1
 
 
    rs("record_date").value = record_date
    rs.update
 
 
 
 
  If val(CboDrawingType.ListIndex) = 3 Then
  
      rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = val(XPTxtVal.text)
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = val(XPTxtID.text)
    rs("kedno").value = val(XPTxtID.text)
        
    rs("opr_type").value = opr_type
    rs("account_name").value = DcboDebitSide1.text
    rs("account_no").value = DcboDebitSide1.BoundText
  '  rs("line_no").value = 1
     rs("line_no").value = Line3
     
          rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtNoteSerial1.text
        rs("Remark").value = " ÕÊÌ·«  „«·Ì… »—ﬁ„ " & TxtNoteSerial1 & "    " & Me.XPMTxtRemarks1
 
 
    rs("record_date").value = record_date
    rs.update
 
 
  End If
  
 
 
 
    rs.Close
End Function
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim DblDif As Double
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then

        '--------------------------------------------------------------------------------
        If Me.CboDrawingType.ListIndex = 1 Then
            'Â–Â ⁄„·Ì…  ÕÊÌ· „‰ Œ“‰… ≈·Ï Œ“‰…
            'ÌÃ» „·«ÕŸ… «‰ «·„” Œœ„ „„ﬂ‰ «‰ ÌﬁÊ„ » ⁄œÌ· «·ﬁÌ„…
            '«· Ï  „  ÕÊÌ·Â« ≈·Ï «·Œ“‰… ÊÌ÷⁄ ﬁÌ„… «ﬁ·
            ' ÊÂ‰« ÌÃ» «· «ﬂœ „‰ «‰ —’Ìœ «·Œ“‰… Ì”„Õ
            DblDif = val(XPTxtVal.text)

            If DblDif > 0 Then
                If CheckBoxAccount(val(Me.DcboBoxTo.BoundText), DblDif, Me.XPDtbTrans.value, False) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
                    Msg = Msg & CHR(13) & "”»» «·Œÿ« ÂÊ «‰ﬂ  Õ«Ê· Õ–› ⁄„·Ì… «· ÕÊÌ· "
                    Msg = Msg & CHR(13) & ""
                    Msg = Msg & CHR(13) & "Ê⁄„·Ì… «·Õ–› Â–Â ”Ê›  Êœ∆ ≈·Ï «‰ ÌﬂÊ‰ "
                    Msg = Msg & CHR(13) & " —’Ìœ «·Œ“‰… ”«·» ›Ï ‰Â«Ì… Â–« «·ÌÊ„ " & DisplayDate(Me.XPDtbTrans.value)
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

        '--------------------------------------------------------------------------------
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (Me.TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

         If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
   Deletepost Me.Name, "Notes", "NoteID", 0, val(dcBranch.BoundText), val(XPTxtID.text), TxtNoteSerial1.text
    
                CuurentLogdata ("D")
                   StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND (  kedno =" & val(XPTxtID.text) & " or  kedno =" & val(Me.TxtNoteID2.text) & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete From NOTES Where NoteId=" & val(TxtNoteID2.text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                '----------------------------------------------------------------------
                rs.delete
                '----------------------------------------------------------------------
              
                rs.MovePrevious

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                   ' GetBoxData
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·”‰œ " & TxtNoteSerial1.text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   ‰Ê⁄ «·”Õ»   " & CboDrawingType & CHR(13) & "  «·›—⁄    " & dcBranch & CHR(13) & "   «·„»·€   " & XPTxtVal & CHR(13) & "    «·»‰ﬂ «·„ÕÊ· „‰…   " & DCBankTo & CHR(13) & "  «·»‰ﬂ «·„ÕÊ· «·Ì…    " & Dcbank & CHR(13) & "    —ﬁ„ «·‘Ìﬂ   " & TxtChequeNumber & CHR(13) & "  «·„” ›Ìœ    " & txtperson & CHR(13) & "      «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "   »‰«¡ ⁄·Ï    " & XPMTxtRemarks
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Serial ¬No " & TxtNoteSerial1.text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "  Type  " & CboDrawingType & CHR(13) & "  Branch    " & dcBranch & CHR(13) & "   Value   " & XPTxtVal & CHR(13) & "   From Bank " & DCBankTo & CHR(13) & "  To Bank   " & Dcbank & CHR(13) & "  Cheque Numbe  " & TxtChequeNumber & CHR(13) & "  To person     " & txtperson & CHR(13) & "  Cheque Due Date  " & DtpChequeDueDate & CHR(13) & "  Based On   " & XPMTxtRemarks
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 14, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 14, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label24.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label24.Caption = "Approved"
                                 End If
                            Label24.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label24.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label24.Caption = "Currently required Approve"
                            End If
                 Label24.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
    End If
RsDetails.Close

End Function
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "  ÕÊÌ·«  „«·Ì…    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ÕÊÌ·«  „«·Ì…    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "      ÕÊÌ·«  „«·Ì…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ÕÊÌ·«  „«·Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ÕÊÌ·«  „«·Ì…    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   ÕÊÌ·«  „«·Ì…    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "     ÕÊÌ·«  „«·Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ÕÊÌ·«  „«·Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "    ÕÊÌ·«  „«·Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ÕÊÌ·«  „«·Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "     ÕÊÌ·«  „«·Ì…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    
        TxtNoteSerial2.text = ""
     
    
End Sub

Private Sub XPMTxtRemarks1_Change()

    XPMTxtRemarks.text = XPMTxtRemarks1.text
End Sub

Private Sub XPTxtVal_Change()
    Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")
    Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text)
End Sub

Private Sub GetBoxData()

    Me.LblBoxName = Me.DcboBox.text
    Me.LblBoxAccount.Caption = get_balanceFromGl(ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText)))

    'Me.LblBoxName = Me.DcboBox.text
    'Me.LblBoxAccount.Caption = GetBoxAccount(Val(Me.DcboBox.BoundText))
End Sub

Private Sub WriteDev()

    '
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.CboDrawingType.ListIndex = 0 Then
            '”Õ» „‰ «·Œ“‰…
 
            Me.DcboDebitSide.BoundText = getBranchCurrentAccount(val(dcBranch.BoundText)) ' «·Ã«—Ì
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
        ElseIf Me.CboDrawingType.ListIndex = 1 Then
            ' ÕÊÌ· „‰ Œ“‰… ·Œ“‰…
            Me.Dcbank.BoundText = ""
            Me.DCBankTo.BoundText = ""
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBoxTo.BoundText))
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
        ElseIf Me.CboDrawingType.ListIndex = 2 Then
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.Dcbank.BoundText))
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DCBankTo.BoundText))
            Me.DcboBox.BoundText = ""
            Me.DcboBoxTo.BoundText = ""
         
        ElseIf Me.CboDrawingType.ListIndex = 3 Then '  ÕÊÌ· »Ì‰ «·›—Ê⁄
     
            Me.DcboDebitSide.BoundText = DCAccounts2.BoundText
            Me.DcboCreditSide.BoundText = DCAccounts1.BoundText
            Me.DcboDebitSide1.BoundText = getBranchCurrentAccount(val(Dcbranch2.BoundText))
            Me.DcboCreditSide1.BoundText = getBranchCurrentAccount(val(dcBranch1.BoundText))
         
            Me.DcboBox.BoundText = ""
            Me.DcboBoxTo.BoundText = ""
            Me.Dcbank.BoundText = ""
            Me.DCBankTo.BoundText = ""
    
        Else
            Me.DcboDebitSide.BoundText = ""
            Me.DcboCreditSide.BoundText = ""
        End If
    End If

End Sub

