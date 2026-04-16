VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStudentsCandidacy 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14085
   Icon            =   "FrmStudentsCandidacy.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10997.22
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   21450.84
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8025
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   14085
      _cx             =   24844
      _cy             =   14155
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
      Align           =   5
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   14130
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   12375
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   44
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   495
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   720
         Width           =   14145
         _cx             =   24950
         _cy             =   873
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
            Height          =   270
            Left            =   165
            MaxLength       =   50
            TabIndex        =   79
            Top             =   120
            Width           =   1050
         End
         Begin VB.TextBox TxtContCode 
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
            Height          =   270
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   77
            Top             =   120
            Width           =   990
         End
         Begin VB.TextBox TxtContNoID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3945
            MaxLength       =   50
            TabIndex        =   4
            Top             =   -480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   11160
            TabIndex        =   0
            Top             =   120
            Width           =   1785
         End
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   270
            Left            =   6900
            TabIndex        =   2
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   476
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   270
            Left            =   8505
            TabIndex        =   1
            Top             =   120
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   476
            _Version        =   393216
            Format          =   93323265
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   120
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   150
            Index           =   11
            Left            =   840
            TabIndex        =   80
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   150
            Index           =   0
            Left            =   2505
            TabIndex        =   59
            Top             =   120
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   240
            Index           =   11
            Left            =   5670
            TabIndex        =   55
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   240
            Index           =   25
            Left            =   9930
            TabIndex        =   54
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   210
            Index           =   4
            Left            =   13185
            TabIndex        =   28
            Top             =   120
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1095
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   6930
         Width           =   14085
         _cx             =   24844
         _cy             =   1931
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
         Align           =   2
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   345
            Left            =   12825
            TabIndex        =   30
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   570
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":6852
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   345
            Left            =   10920
            TabIndex        =   31
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   570
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":D0B4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   345
            Left            =   9585
            TabIndex        =   14
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   570
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":13916
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   345
            Left            =   7680
            TabIndex        =   32
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   570
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":13CB0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   345
            Left            =   6225
            TabIndex        =   33
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   570
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":1404A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   450
            Left            =   3630
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   570
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   794
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":145E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   345
            Left            =   5040
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   570
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":1AE46
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   345
            Left            =   240
            TabIndex        =   36
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   570
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":1B1E0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   9015
            TabIndex        =   37
            Top             =   90
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   300
            Left            =   6195
            TabIndex        =   52
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            ButtonStyle     =   1
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":1B57A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   450
            Left            =   1425
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   570
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   794
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··»—Ìœ «·«·ﬂ —Ê‰Ì"
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":21DDC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   150
            Left            =   315
            TabIndex        =   42
            Top             =   255
            Width           =   690
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   150
            Left            =   2550
            TabIndex        =   41
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   150
            Index           =   1
            Left            =   1155
            TabIndex        =   40
            Top             =   255
            Width           =   1230
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   150
            Index           =   0
            Left            =   3465
            TabIndex        =   39
            Top             =   255
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   14
            Left            =   13125
            TabIndex        =   38
            Top             =   90
            Width           =   1275
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   0
         Width           =   14130
         _cx             =   24924
         _cy             =   1376
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
         BackColor       =   16777215
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
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   135
            TabIndex        =   47
            Top             =   240
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":2863E
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   720
            TabIndex        =   48
            Top             =   240
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":289D8
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1395
            TabIndex        =   49
            Top             =   240
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":28D72
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2070
            TabIndex        =   50
            Top             =   240
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":2910C
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13005
            Picture         =   "FrmStudentsCandidacy.frx":294A6
            Stretch         =   -1  'True
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· —‘ÌÕ  ··‘—ﬂ« "
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
            Left            =   5910
            TabIndex        =   51
            Top             =   240
            Width           =   4935
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1530
         Left            =   0
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1275
         Width           =   14145
         _cx             =   24950
         _cy             =   2699
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
         Begin VB.TextBox TxtCode 
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
            Height          =   375
            Left            =   10635
            MaxLength       =   50
            TabIndex        =   5
            Top             =   180
            Width           =   1545
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
            Height          =   855
            Left            =   1755
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   615
            Width           =   2190
         End
         Begin VB.TextBox TxtTrainingID 
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
            Height          =   360
            Left            =   7380
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1470
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox TxtNoStudRemain 
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
            Height          =   360
            Left            =   7590
            MaxLength       =   50
            TabIndex        =   10
            Top             =   615
            Width           =   1065
         End
         Begin VB.TextBox TxtNoStudAccept 
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
            Height          =   330
            Left            =   10635
            MaxLength       =   50
            TabIndex        =   9
            Top             =   615
            Width           =   1545
         End
         Begin VB.TextBox TxtNoStusCandid 
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
            Height          =   360
            Left            =   4650
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   8
            Top             =   615
            Width           =   1050
         End
         Begin VB.TextBox TxtNoStudCon 
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
            Height          =   375
            Left            =   120
            MaxLength       =   50
            TabIndex        =   7
            Top             =   180
            Width           =   1980
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   855
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   615
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1508
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":2A8AB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbCompany 
            Height          =   315
            Left            =   4650
            TabIndex        =   6
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   180
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal DataFomH 
            Height          =   330
            Left            =   9120
            TabIndex        =   71
            Top             =   1095
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker DataFom 
            Height          =   345
            Left            =   10635
            TabIndex        =   72
            Top             =   1095
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   93323265
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal DateToH 
            Height          =   330
            Left            =   4650
            TabIndex        =   74
            Top             =   1095
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker DateTo 
            Height          =   345
            Left            =   6210
            TabIndex        =   75
            Top             =   1095
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   93323265
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   210
            Index           =   1
            Left            =   7905
            TabIndex        =   76
            Top             =   1095
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   210
            Index           =   0
            Left            =   12045
            TabIndex        =   73
            Top             =   1095
            Width           =   1845
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ…"
            Height          =   210
            Index           =   5
            Left            =   12090
            TabIndex        =   65
            Top             =   195
            Width           =   1860
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ ÿ·» «· œ—Ì»"
            Height          =   195
            Index           =   3
            Left            =   8805
            TabIndex        =   64
            Top             =   1560
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   4
            Left            =   3915
            TabIndex        =   63
            Top             =   825
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·„—‘ÕÌ‰ «·„ »ﬁÌ‰"
            Height          =   195
            Index           =   14
            Left            =   8700
            TabIndex        =   62
            Top             =   600
            Width           =   1845
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·„—‘ÕÌ‰ «·„ﬁ»Ê·Ì‰"
            Height          =   195
            Index           =   10
            Left            =   12090
            TabIndex        =   61
            Top             =   600
            Width           =   1860
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ  «·„—‘ÕÌ‰"
            Height          =   210
            Index           =   8
            Left            =   5745
            TabIndex        =   60
            Top             =   600
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·„—‘ÕÌ‰ ›Ì «·⁄ﬁœ"
            Height          =   225
            Index           =   1
            Left            =   2280
            TabIndex        =   56
            Top             =   195
            Width           =   1770
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   4005
         Left            =   0
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2850
         Width           =   14145
         _cx             =   24950
         _cy             =   7064
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
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   3675
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   13995
            _cx             =   24686
            _cy             =   6482
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16776960
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmStudentsCandidacy.frx":3110D
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   0
            Left            =   12990
            TabIndex        =   67
            Top             =   3690
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmStudentsCandidacy.frx":3136C
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   10725
            TabIndex        =   68
            Top             =   3690
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ·"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmStudentsCandidacy.frx":31906
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3645
            Left            =   0
            TabIndex        =   81
            Top             =   0
            Width           =   13995
            _cx             =   24686
            _cy             =   6429
            Appearance      =   2
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmStudentsCandidacy.frx":31EA0
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«··Ê‰ «·«Õ„— Ìœ· ⁄·Ï «‰ «·„—‘Õ  „  —‘ÌÕ… ·⁄ﬁœ «Œ—"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   9
            Left            =   765
            TabIndex        =   78
            Top             =   3690
            Width           =   3795
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   7
            Left            =   4395
            TabIndex        =   70
            Top             =   3690
            Width           =   1830
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ  «·„—‘ÕÌ‰"
            Height          =   195
            Index           =   6
            Left            =   6300
            TabIndex        =   69
            Top             =   3690
            Width           =   1815
         End
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      TabIndex        =   19
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmStudentsCandidacy.frx":32205
      Left            =   15480
      List            =   "FrmStudentsCandidacy.frx":32215
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   20
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
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
            Picture         =   "FrmStudentsCandidacy.frx":3222E
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":325C8
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":32962
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":32CFC
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":33096
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":33430
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":337CA
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStudentsCandidacy.frx":33D64
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmStudentsCandidacy.frx":340FE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      ButtonImage     =   "FrmStudentsCandidacy.frx":3A960
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
      ButtonImage     =   "FrmStudentsCandidacy.frx":411C2
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
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
      Left            =   15480
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmStudentsCandidacy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim II As Long


Private Sub Cmd_Click(Index As Integer)
Dim i As Integer
Dim k As Integer
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow

Case 1
k = Fg1.Rows - 1
For i = 1 To Fg1.Rows - 1

If RemoveGridRowAll(k) = True Then
Exit Sub
End If
If k <= 0 Then Exit Sub
Fg1.RemoveItem k
k = k - 1
Next i
End Select
End If
End Sub
Function RemoveGridRowAll(Optional i As Integer) As Boolean
    With Me.Fg1
        If i = 0 Then Exit Function
        If CheckAccespt(val(Fg1.TextMatrix(i, Fg1.ColIndex("ID")))) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› «Ê «· ⁄œÌ· „— »ÿ »„Ê«›ﬁ… «· —‘ÌÕ"
        Else
        MsgBox "Can Not Delete or edit linked to Accept Nomination"
        End If
        RemoveGridRowAll = True
        Exit Function
        Else
        RemoveGridRowAll = False
        End If
    End With
    ReLineGrid
End Function

Private Sub RemoveGridRow()
    With Me.Fg1

        If .Row <= 0 Then Exit Sub
        If CheckAccespt(val(.TextMatrix(.Row, .ColIndex("ID")))) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ– „— »ÿ «· —‘ÌÕ"
        Else
        MsgBox "Can Not Delete linked to Accept Nomination"
        End If
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
    ReLineGrid
End Sub
Sub ReLineGrid()
Dim i As Integer
Dim Conter As Integer
Conter = 0
With Fg1
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Name")) <> "" Then
Conter = Conter + 1
.TextMatrix(i, .ColIndex("Serial")) = Conter
End If
Next i
TxtNoStusCandid.Text = Conter
Label1(7).Caption = Conter
End With
End Sub

Private Sub DataFom_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(DataFom.value) Then
         DataFomH.value = ToHijriDate(DataFom.value)
End If
End If
End Sub

Private Sub DataFomH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 DataFom.value = ToGregorianDate(DataFomH.value)
End If
End Sub

Private Sub DateTo_Change()
If Me.TxtModFlg.Text <> "R" Then
If Not IsNull(DateTo.value) Then
         DateToH.value = ToHijriDate(DateTo.value)
End If
End If
End Sub

Private Sub DateToH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 DateTo.value = ToGregorianDate(DateToH.value)
End If
End Sub

Private Sub DcbCompany_Change()
DcbCompany_Click (0)
End Sub
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblStuCandidacy", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Private Sub DcbCompany_Click(Area As Integer)
  If val(DcbCompany.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany.BoundText, EmpCode
    Me.TxtCode.Text = EmpCode
End Sub


Sub filgrid1()
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim i As Integer
Dim k As Integer
sql = "SELECT     dbo.TblStudentQualification.Name AS QName, dbo.TblStudentQualification.NameE AS QNameE, dbo.TblTrainingRequest.*"
sql = sql & " FROM         dbo.TblTrainingRequest LEFT OUTER JOIN"
sql = sql & "                     dbo.TblStudentQualification ON dbo.TblTrainingRequest.QualiID = dbo.TblStudentQualification.ID"
sql = sql & " Where (dbo.TblTrainingRequest.TypeTrain = 1) and (dbo.TblTrainingRequest.FlagAccept is null) "
sql = sql & " and  dbo.TblTrainingRequest.id not in(SELECT     dbo.TblStuCandidacyDet.TrainingID"
sql = sql & " FROM         dbo.TblStuCandidacy LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacyDet ON dbo.TblStuCandidacy.ID = dbo.TblStuCandidacyDet.StudCandID"
sql = sql & " WHERE     (dbo.TblStuCandidacy.ContNoID = " & val(TxtContNoID.Text) & ") ) "
sql = sql & "  and   (dbo.TblTrainingRequest.BranchID=0 or dbo.TblTrainingRequest.BranchID is null or  dbo.TblTrainingRequest.BranchID in(" & Current_branchSql & "))"
    
If Not IsNull(DataFom.value) Then
sql = sql & " and dbo.TblTrainingRequest.RecordDate>=" & SQLDate(DataFom.value, True) & ""
End If
If Not IsNull(DateTo.value) Then
sql = sql & " and dbo.TblTrainingRequest.RecordDate<=" & SQLDate(DateTo.value, True) & ""
End If
If SystemOptions.UserInterface = ArabicInterface Then
sql = sql & " order by dbo.TblTrainingRequest.Name"
Else
sql = sql & " order by dbo.TblTrainingRequest.NameE"
End If
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 Fg1.Clear flexClearScrollable, flexClearEverything
     Fg1.Rows = 1
If Rs2.RecordCount > 0 Then
With Fg1
k = .Rows
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("TrainingID")) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
If Check(val(TxtContNoID.Text), .TextMatrix(i, .ColIndex("TrainingID"))) = True Then
.Cell(flexcpBackColor, i, 1, i, 15) = &HFF&
End If
.TextMatrix(i, .ColIndex("QualiID")) = IIf(IsNull(Rs2("QualiID").value), 0, Rs2("QualiID").value)
.TextMatrix(i, .ColIndex("SexID")) = IIf(IsNull(Rs2("SexID").value), -1, Rs2("SexID").value)
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs2("FullCode").value), "", Rs2("FullCode").value)
.TextMatrix(i, .ColIndex("Jeha")) = IIf(IsNull(Rs2("Jeha").value), "", Rs2("Jeha").value)
.TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(Rs2("UQama").value), "", Rs2("UQama").value)
.TextMatrix(i, .ColIndex("Experience")) = IIf(IsNull(Rs2("Experience").value), "", Rs2("Experience").value)
.TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(Rs2("Phone").value), "", Rs2("Phone").value)
.TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(Rs2("Email").value), "", Rs2("Email").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs2("RecordDate").value), Date, Rs2("RecordDate").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs2("RecordDateH").value), "", Rs2("RecordDateH").value)
.TextMatrix(i, .ColIndex("Remarks")) = Me.TxtRemarks.Text
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
.TextMatrix(i, .ColIndex("QName")) = IIf(IsNull(Rs2("QName").value), "", Rs2("QName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
.TextMatrix(i, .ColIndex("QName")) = IIf(IsNull(Rs2("QNameE").value), "", Rs2("QNameE").value)
End If
Rs2.MoveNext
Next i
End With
End If
End Sub
Function Check(Optional ID As Double, Optional TrainingID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblStuCandidacyDet.TrainingID"
sql = sql & " FROM         dbo.TblStuCandidacy LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStuCandidacyDet ON dbo.TblStuCandidacy.ID = dbo.TblStuCandidacyDet.StudCandID"
sql = sql & " Where (dbo.TblStuCandidacy.ContNoID <> " & ID & ") and dbo.TblStuCandidacyDet.TrainingID=" & TrainingID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Check = True
Else
Check = False
End If
End Function
Function CheckAccespt(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "Select *from TblStuCandidacyDet where AccptedID=1 and id=" & ID & " "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheckAccespt = True
Else
CheckAccespt = False
End If
End Function

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    If mdifrmmain.xyz = True Then
    Fg1.Visible = False
   Grid.Visible = True
    Else
    Fg1.Visible = True
    Grid.Visible = False
    End If
    conection = "select * from  TblStuCandidacy  "
    conection = conection & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    conection = conection & " Order By ID"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   Dcombos.GetBranches Me.DcbBranch
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany
   
    BtnLast_Click
    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
   FiLLTXT
ErrTrap:
End Sub
Function GetNoAccept() As Double
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = " SELECT     ContNoID, SUM(NoStudACNow) AS SmNoStudACNow"
sql = sql & " From dbo.TblStuCandidacyAccept"
sql = sql & " Where   (ContNoID = " & val(TxtContNoID.Text) & ")"

sql = sql & " GROUP BY ContNoID"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetNoAccept = IIf(IsNull(Rs4("SmNoStudACNow").value), 0, Rs4("SmNoStudACNow").value)
Else
GetNoAccept = 0
End If
End Function
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    Dim i As Integer
      If Me.TxtModFlg.Text = "E" Then
  For i = 1 To Fg1.Rows - 1
  If RemoveGridRowAll(i) = True Then
  Exit Sub
  End If
  Next i
 Cn.Execute " delete from TblStuCandidacyDet where StudCandID=" & val(TxtSerial1.Text) & "  "
 End If
 
   RsSavRec.Fields("ContCode").value = TxtContCode.Text
   RsSavRec.Fields("RecordDateH").value = RecordDateH.value
   RsSavRec.Fields("RecordDate").value = RecordDate.value
   RsSavRec.Fields("DataFomH").value = DataFomH.value
   RsSavRec.Fields("DataFom").value = DataFom.value
   RsSavRec.Fields("DateToH").value = DateToH.value
   RsSavRec.Fields("DateTo").value = DateTo.value
   RsSavRec.Fields("CompID").value = val(DcbCompany.BoundText)
   RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("Remarks").value = TxtRemarks.Text
   RsSavRec.Fields("TrainingID").value = val(TxtTrainingID.Text)
   RsSavRec.Fields("ContNoID").value = val(TxtContNoID.Text)
   RsSavRec.Fields("NoStudCon").value = val(TxtNoStudCon.Text)
   RsSavRec.Fields("NoStudRemain").value = val(TxtNoStudRemain.Text)
   RsSavRec.Fields("NoStudAccept").value = val(TxtNoStudAccept.Text)
   RsSavRec.Fields("NoStusCandid").value = val(TxtNoStusCandid.Text)
   RsSavRec.update
  ''//////////////////////////

  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblStuCandidacyDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim str2 As String
    With Me.Fg1
       For i = .FixedRows To .Rows - 1
       If (.TextMatrix(i, .ColIndex("Name"))) <> "" Then
       RsDevsub.AddNew
                RsDevsub("StudCandID").value = val(Me.TxtSerial1.Text)
                RsDevsub("FullCode").value = IIf((.TextMatrix(i, .ColIndex("FullCode"))) = "", Null, .TextMatrix(i, .ColIndex("FullCode")))
                RsDevsub("TrainingID").value = IIf((.TextMatrix(i, .ColIndex("TrainingID"))) = "", Null, val(.TextMatrix(i, .ColIndex("TrainingID"))))
                RsDevsub("Name").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, .TextMatrix(i, .ColIndex("Name")))
                RsDevsub("QualiID").value = IIf((.TextMatrix(i, .ColIndex("QualiID"))) = "", Null, val(.TextMatrix(i, .ColIndex("QualiID"))))
                RsDevsub("Jeha").value = IIf((.TextMatrix(i, .ColIndex("Jeha"))) = "", Null, (.TextMatrix(i, .ColIndex("Jeha"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, .TextMatrix(i, .ColIndex("Remarks")))
                RsDevsub("Experience").value = IIf((.TextMatrix(i, .ColIndex("Experience"))) = "", Null, (.TextMatrix(i, .ColIndex("Experience"))))
                RsDevsub("Phone").value = IIf((.TextMatrix(i, .ColIndex("Phone"))) = "", Null, .TextMatrix(i, .ColIndex("Phone")))
                RsDevsub("Email").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, .TextMatrix(i, .ColIndex("Email")))
                RsDevsub("UQama").value = IIf((.TextMatrix(i, .ColIndex("UQama"))) = "", Null, .TextMatrix(i, .ColIndex("UQama")))
                RsDevsub("SexID").value = IIf((.TextMatrix(i, .ColIndex("SexID"))) = "", Null, val(.TextMatrix(i, .ColIndex("SexID"))))
                RsDevsub("RecordDateH").value = IIf((.TextMatrix(i, .ColIndex("RecordDateH"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDateH"))))
                RsDevsub("RecordDate").value = IIf((.TextMatrix(i, .ColIndex("RecordDate"))) = "", Null, (.TextMatrix(i, .ColIndex("RecordDate"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
   
    Dim Msg As String
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
Sub FullGri()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Dim sql As String
    Fg1.Clear flexClearScrollable, flexClearEverything
    Fg1.Rows = 2
sql = "SELECT     dbo.TblStuCandidacyDet.ID, dbo.TblStuCandidacyDet.StudCandID, dbo.TblStuCandidacyDet.Name, dbo.TblStuCandidacyDet.Jeha, dbo.TblStuCandidacyDet.Remarks, "
sql = sql & "                      dbo.TblStuCandidacyDet.Experience, dbo.TblStuCandidacyDet.Phone, dbo.TblStuCandidacyDet.Email, dbo.TblStuCandidacyDet.SexID,"
sql = sql & "                      dbo.TblStuCandidacyDet.UQama, dbo.TblStuCandidacyDet.QualiID, dbo.TblStudentQualification.Name AS QName,"
sql = sql & "                      dbo.TblStudentQualification.NameE AS QNameE,dbo.TblStuCandidacyDet.TrainingID,dbo.TblStuCandidacyDet.FullCode ,"
sql = sql & "                      dbo.TblStuCandidacyDet.RecordDate,dbo.TblStuCandidacyDet.RecordDateH"
sql = sql & " FROM         dbo.TblStuCandidacyDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStudentQualification ON dbo.TblStuCandidacyDet.QualiID = dbo.TblStudentQualification.ID"
sql = sql & " Where (dbo.TblStuCandidacyDet.StudCandID =" & val(TxtSerial1.Text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Fg1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
.TextMatrix(i, .ColIndex("FullCode")) = IIf(IsNull(Rs3("FullCode").value), "", Rs3("FullCode").value)
.TextMatrix(i, .ColIndex("TrainingID")) = IIf(IsNull(Rs3("TrainingID").value), 0, Rs3("TrainingID").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(i, .ColIndex("QualiID")) = IIf(IsNull(Rs3("QualiID").value), 0, Rs3("QualiID").value)
.TextMatrix(i, .ColIndex("Jeha")) = IIf(IsNull(Rs3("Jeha").value), "", Rs3("Jeha").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs3("Remarks").value), "", Rs3("Remarks").value)
.TextMatrix(i, .ColIndex("Experience")) = IIf(IsNull(Rs3("Experience").value), "", Rs3("Experience").value)
.TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(Rs3("Phone").value), "", Rs3("Phone").value)
.TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(Rs3("Email").value), "", Rs3("Email").value)
.TextMatrix(i, .ColIndex("SexID")) = IIf(IsNull(Rs3("SexID").value), -1, Rs3("SexID").value)
.TextMatrix(i, .ColIndex("UQama")) = IIf(IsNull(Rs3("UQama").value), "", Rs3("UQama").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs3("RecordDateH").value), "", Rs3("RecordDateH").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs3("RecordDate").value), "", Rs3("RecordDate").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("QName")) = IIf(IsNull(Rs3("QName").value), "", Rs3("QName").value)
Else
.TextMatrix(i, .ColIndex("QName")) = IIf(IsNull(Rs3("QNameE").value), "", Rs3("QNameE").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()

   On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
    
    TxtContCode.Text = IIf(IsNull(RsSavRec.Fields("ContCode").value), "", RsSavRec.Fields("ContCode").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbCompany.BoundText = IIf(IsNull(RsSavRec.Fields("CompID").value), "", RsSavRec.Fields("CompID").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    Me.TxtTrainingID.Text = IIf(IsNull(RsSavRec.Fields("TrainingID").value), "", RsSavRec.Fields("TrainingID").value)
    TxtContNoID.Text = IIf(IsNull(RsSavRec.Fields("ContNoID").value), "", RsSavRec.Fields("ContNoID").value)
    TxtNoStudCon.Text = IIf(IsNull(RsSavRec.Fields("NoStudCon").value), 0, RsSavRec.Fields("NoStudCon").value)
     TxtNoStudAccept.Text = IIf(IsNull(RsSavRec.Fields("NoStudAccept").value), 0, RsSavRec.Fields("NoStudAccept").value)
     Me.TxtNoStudRemain.Text = IIf(IsNull(RsSavRec.Fields("NoStudRemain").value), 0, RsSavRec.Fields("NoStudRemain").value)
    TxtNoStusCandid.Text = IIf(IsNull(RsSavRec.Fields("NoStusCandid").value), 0, RsSavRec.Fields("NoStusCandid").value)
    Label1(7).Caption = IIf(IsNull(RsSavRec.Fields("NoStusCandid").value), 0, RsSavRec.Fields("NoStusCandid").value)
    DataFomH.value = IIf(IsNull(RsSavRec.Fields("DataFomh").value), ToHijriDate(Date), RsSavRec.Fields("DataFomh").value)
    DataFom.value = IIf(IsNull(RsSavRec.Fields("DataFom").value), Date, RsSavRec.Fields("DataFom").value)
    DateToH.value = IIf(IsNull(RsSavRec.Fields("DateToH").value), ToHijriDate(Date), RsSavRec.Fields("DateToH").value)
    DateTo.value = IIf(IsNull(RsSavRec.Fields("DateTo").value), Date, RsSavRec.Fields("DateTo").value)
    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    If val(DcbBranch.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
    Else
    MsgBox "Please Select Branch"
    End If
    DcbBranch.SetFocus
    Exit Sub
    End If
    If val(TxtContNoID.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·⁄ﬁœ "
    Else
    MsgBox "Please Eneter Contract No."
    End If
    TxtContNoID.SetFocus
    Exit Sub
    End If
    'If val(TxtTrainingID.text) = 0 Then
   ' If SystemOptions.UserInterface = ArabicInterface Then
   ' MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ ÿ·» «· œ—Ì» "
   ' Else
   ' MsgBox "Please Eneter Training  No."
   ' End If
   ' TxtTrainingID.SetFocus
   ' Exit Sub
   ' End If
    
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Function CheckRepeat() As Boolean
Dim i As Integer
With Fg1
For i = 1 To .Rows - 1
If val(TxtTrainingID.Text) = val(.TextMatrix(i, .ColIndex("TrainingID"))) Then
CheckRepeat = True
Exit Function
End If
Next i
End With
CheckRepeat = False
End Function
Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
If Me.TxtModFlg.Text <> "N" Then
If val(TxtNoStusCandid.Text) + 1 > val(TxtNoStudRemain.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄œœ «·ÿ·«» «·„—‘ÕÌ‰ «ﬂ»— „‰ ⁄œœ «·ÿ·«» «·„ »ﬁÌ‰"
Else
MsgBox "Can not be larger than the remaining number"
End If
Exit Sub
End If
End If
    If (TxtContCode.Text) = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·⁄ﬁœ "
    Else
    MsgBox "Please Eneter Contract No."
    End If
    TxtContCode.SetFocus
    Exit Sub
    End If
     If val(DcbCompany.BoundText) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·‘—ﬂ… "
    Else
    MsgBox "Please Select Company"
    End If
    DcbCompany.SetFocus
    Exit Sub
    End If
    
 '   If val(TxtTrainingID.text) = 0 Then
  '  If SystemOptions.UserInterface = ArabicInterface Then
   ' MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ ÿ·» «· œ—Ì» "
    'Else
  '  MsgBox "Please Eneter Training  No."
   ' End If
   ' TxtTrainingID.SetFocus
   ' Exit Sub
   ' End If
'  If CheckRepeat() = True Then
 ' If SystemOptions.UserInterface = ArabicInterface Then
  'MsgBox "·«Ì„ﬂ‰ «· ﬂ—«—"
'  Else
 ' MsgBox "Can Not Repetition"
 ' End If
 ' Exit Sub
 ' End If
 
filgrid1
ReLineGrid
End If
End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
ShowAttachments TxtSerial1.Text, "04092016001"
ErrTrap:
End Sub

Private Sub ISButton8_Click()
FrmSearStudent.inde = 5
Load FrmSearStudent
FrmSearStudent.show vbModal
End Sub

Private Sub RecordDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(RecordDate.value)
End If
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
 RecordDate.value = ToGregorianDate(RecordDateH.value)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim RsDev As ADODB.Recordset
      Set RsDev = New ADODB.Recordset
      
    If KeyAscii = vbKeyReturn Then

   Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2

StrSQL = " SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
 StrSQL = StrSQL & "                      dbo.TbVisaDeti.NotionalID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
 StrSQL = StrSQL & "                      dbo.TbVisaDeti.[count] , dbo.TbVisaDeti.Type, dbo.TbVisaDeti.Place , dbo.TbVisaDeti.Price"
StrSQL = StrSQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCountriesGovernments RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TbVisaDeti ON dbo.TblCountriesGovernments.GovernmentID = dbo.TbVisaDeti.CityID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Nationality ON dbo.TbVisaDeti.NotionalID = dbo.Nationality.id ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TbVisaDeti.JobID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  Where (dbo.TbVisaDeti.VisaID = " & val(Text1.Text) & ") And (dbo.TbVisaDeti.Type = 0)"

     
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
Dim i As Integer
    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Job")) = IIf(IsNull(RsDev("JobTypeNamee").value), "", RsDev("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("Notional")) = IIf(IsNull(RsDev("namee").value), "", RsDev("namee").value)
                End If
                .TextMatrix(i, .ColIndex("HododNo")) = IIf(IsNull(RsDev("HododNo").value), "", RsDev("HododNo").value)
                .TextMatrix(i, .ColIndex("City")) = IIf(IsNull(RsDev("GovernmentName").value), "", RsDev("GovernmentName").value)
                .TextMatrix(i, .ColIndex("NotionalID")) = IIf(IsNull(RsDev("NotionalID").value), "", RsDev("NotionalID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
                .TextMatrix(i, .ColIndex("CityID")) = IIf(IsNull(RsDev("CityID").value), "", RsDev("CityID").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), "", RsDev("Price").value)
            
            
                RsDev.MoveNext
            Next i
 
        End With

    End If


    End If
    
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtCode.Text, EmpID
        DcbCompany.BoundText = EmpID
    End If
End Sub

Private Sub TxtContCode_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtContNoID = GetContID(TxtContCode.Text)
TxtContNoID_Change
End If
End Sub

Private Sub TxtContCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 401
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub TxtContNoID_Change()
Dim CompID As Double
Dim NoStud As Double
If Me.TxtModFlg.Text <> "R" Then
If val(Me.TxtContNoID.Text) <> 0 Then
GetContStudentInformation val(Me.TxtContNoID.Text), CompID, NoStud
DcbCompany.BoundText = CompID
TxtNoStudCon.Text = NoStud
TxtNoStudAccept.Text = GetNoAccept()
TxtNoStudRemain.Text = val(TxtNoStudCon.Text) - val(TxtNoStudAccept.Text)
End If
End If
End Sub
Function GetContID(Optional Fullcode As String) As Double
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
Dim sql As String
sql = " SELECT id From dbo.TblContrStudent where  Fullcode='" & Fullcode & "'"
Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs1.RecordCount > 0 Then
GetContID = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
Else
GetContID = 0
End If
End Function
Private Sub TxtContNoID_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtContNoID.Text, 0)
End Sub

Private Sub TxtNoStudAccept_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoStudRemain.Text = val(TxtNoStudCon.Text) - val(TxtNoStudAccept.Text)
End If
End Sub

Private Sub TxtNoStudCon_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtNoStudRemain.Text = val(TxtNoStudCon.Text) - val(TxtNoStudAccept.Text)
End If
End Sub

Private Sub TxtNoStusCandid_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoStusCandid.Text, 0)
End Sub


' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim sql As String
    Dim i As Integer
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim ID As Double
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
 
  For i = 1 To Fg1.Rows - 1
  If RemoveGridRowAll(i) = True Then
  Exit Sub
  End If
  Next i
  
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
         Cn.Execute "Delete from TblStuCandidacyDet where StudCandID=" & val(TxtSerial1.Text) & " "
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            Fg1.Clear flexClearScrollable, flexClearEverything
                  Fg1.Rows = 1
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
     Label1(7).Caption = 0
     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub

' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    'XPDtbTrans.Enabled = True
      '  Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
   ' XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
  ' XPDtbTrans.Enabled = True
  '     Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
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
            Msg = Msg & "By another user on the network " & CHR(13)
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
        FindRec val(TxtSerial1.Text)
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
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
    If TxtSerial1.Text <> "" Then
    Fg1.Rows = Fg1.Rows + 1
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
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
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
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
    clear_all Me
    TxtModFlg.Text = "N"
    Me.DcbBranch.BoundText = Current_branch
      Fg1.Clear flexClearScrollable, flexClearEverything
     Fg1.Rows = 1
    Me.DCboUserName.BoundText = user_id
 Label1(7).Caption = 0
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
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
            Msg = Msg & "By another user on the network " & CHR(13)
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
        FindRec val(TxtSerial1.Text)
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
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

'Information for camand
'++++++++++++++++++++++++++++++++++++++
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
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
On Error GoTo ErrTrap
   ' form name
  Label1(2).Caption = "Students Nomination To Companies "
lbl(4).Caption = "No"
lbl(25).Caption = "Date"
lbl(11).Caption = "Branch"
ISButton3.Caption = "Attachments"
Label1(0).Caption = "Contract No."
Label1(5).Caption = "Companies"
Label1(1).Caption = "No. Student "
Label1(8).Caption = "No. Nominees"
Label1(14).Caption = "No. Remaining"
Label1(10).Caption = "No. Accepted"
Label1(4).Caption = "Remarks"
ISButton2.Caption = "Add"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete All"
Label1(6).Caption = "No. Nominees"
lbl(14).Caption = "By"
ISButton4.Caption = "Send an Email"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
   ' C1Tab1.Caption = "Data"

    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    lbl(0).Caption = "From Date"
    lbl(1).Caption = "To Date"
Label1(3).Caption = "No.Training"
  
  With Fg1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("FullCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Student Name"
  .TextMatrix(0, .ColIndex("RecordDateH")) = "Training Date"
  .TextMatrix(0, .ColIndex("RecordDate")) = "Training Date"
  .TextMatrix(0, .ColIndex("TrainingID")) = "Training No."
  .TextMatrix(0, .ColIndex("QName")) = "Qualification"
  .TextMatrix(0, .ColIndex("Jeha")) = "Graduation From"
  .TextMatrix(0, .ColIndex("Experience")) = "Experiences"
  .TextMatrix(0, .ColIndex("Phone")) = "Phone"
  .TextMatrix(0, .ColIndex("Email")) = "Email"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
  End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblStuCandidacy"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub TxtTrainingID_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTrainingID.Text, 0)
End Sub
