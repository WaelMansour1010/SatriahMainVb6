VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmOrderUpload 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmOrderUpload.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   8940
      Left            =   0
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   14235
      _cx             =   25109
      _cy             =   15769
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   7545
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1425
         Width           =   14295
         _cx             =   25215
         _cy             =   13309
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   510
            Left            =   0
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   7020
            Width           =   14295
            _cx             =   25215
            _cy             =   900
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
            Begin ImpulseButton.ISButton btnNew 
               Height          =   330
               Left            =   12060
               TabIndex        =   73
               ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
               Top             =   90
               Width           =   1245
               _ExtentX        =   2196
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
               ButtonImage     =   "FrmOrderUpload.frx":6852
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   8160
               TabIndex        =   74
               ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   90
               Width           =   1350
               _ExtentX        =   2381
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
               ButtonImage     =   "FrmOrderUpload.frx":D0B4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   10470
               TabIndex        =   75
               ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
               Top             =   90
               Width           =   1245
               _ExtentX        =   2196
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
               ButtonImage     =   "FrmOrderUpload.frx":D44E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   6450
               TabIndex        =   76
               ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
               Top             =   90
               Width           =   1485
               _ExtentX        =   2619
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
               ButtonImage     =   "FrmOrderUpload.frx":13CB0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   4755
               TabIndex        =   77
               ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
               Top             =   90
               Width           =   1350
               _ExtentX        =   2381
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
               ButtonImage     =   "FrmOrderUpload.frx":1404A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   75
               TabIndex        =   78
               ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
               Top             =   90
               Width           =   810
               _ExtentX        =   1429
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
               ButtonImage     =   "FrmOrderUpload.frx":145E4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton5 
               Height          =   330
               Left            =   3525
               TabIndex        =   79
               TabStop         =   0   'False
               ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
               Top             =   120
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   582
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
               ButtonImage     =   "FrmOrderUpload.frx":1497E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton8 
               Height          =   330
               Left            =   2370
               TabIndex        =   80
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   90
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
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
               ButtonImage     =   "FrmOrderUpload.frx":1B1E0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton2 
               CausesValidation=   0   'False
               Height          =   405
               Left            =   1425
               TabIndex        =   130
               Top             =   120
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   714
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
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   450
            Left            =   0
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   6585
            Width           =   14295
            _cx             =   25215
            _cy             =   794
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
            Begin MSDataListLib.DataCombo DCboUserName 
               Height          =   315
               Left            =   7680
               TabIndex        =   66
               Top             =   90
               Width           =   4680
               _ExtentX        =   8255
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   270
               Index           =   0
               Left            =   3885
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   90
               Width           =   990
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   270
               Index           =   1
               Left            =   1425
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   90
               Width           =   990
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2565
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   90
               Width           =   1170
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   90
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Õ—— »Ê«”ÿ…  "
               Height          =   270
               Index           =   8
               Left            =   13155
               TabIndex        =   67
               Top             =   90
               Width           =   915
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic12 
            Height          =   1320
            Left            =   -90
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   -45
            Width           =   14295
            _cx             =   25215
            _cy             =   2328
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
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   11715
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   120
               Width           =   1395
            End
            Begin VB.TextBox TxtNationality 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   975
               Width           =   4935
            End
            Begin VB.TextBox TxtIDNo 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   615
               Width           =   4935
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   11715
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   600
               Width           =   1395
            End
            Begin VB.TextBox TxtLeaderName 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   975
               Width           =   6750
            End
            Begin XtremeSuiteControls.RadioButton ChDrievType 
               Height          =   255
               Index           =   0
               Left            =   12885
               TabIndex        =   3
               Top             =   615
               Width           =   1230
               _Version        =   786432
               _ExtentX        =   2170
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "œ«Œ·Ì"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton ChDrievType 
               Height          =   210
               Index           =   1
               Left            =   13095
               TabIndex        =   7
               Top             =   990
               Width           =   1020
               _Version        =   786432
               _ExtentX        =   1799
               _ExtentY        =   370
               _StockProps     =   79
               Caption         =   "Œ«—ÃÌ"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmployee 
               Height          =   315
               Left            =   6360
               TabIndex        =   5
               Top             =   615
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmpSuper 
               Height          =   315
               Left            =   6360
               TabIndex        =   126
               Top             =   120
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   556
               _Version        =   393216
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„—«ﬁ»"
               Height          =   255
               Index           =   23
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   120
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ã‰”Ì…"
               Height          =   225
               Index           =   5
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   975
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÂÊÌ…"
               Height          =   270
               Index           =   1
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   615
               Width           =   1050
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   480
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   255
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·”«∆ﬁ"
               Height          =   255
               Index           =   29
               Left            =   13170
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   360
               Width           =   1050
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   1365
            Left            =   0
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1185
            Width           =   14295
            _cx             =   25215
            _cy             =   2408
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
            Begin VB.TextBox TxtTotal 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   86
               Top             =   1005
               Width           =   2055
            End
            Begin VB.TextBox TxtPartPrice 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3045
               TabIndex        =   84
               Top             =   1005
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4665
               TabIndex        =   82
               Top             =   1005
               Width           =   1560
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   11265
               TabIndex        =   13
               Top             =   675
               Width           =   1065
            End
            Begin XtremeSuiteControls.RadioButton ChCarType 
               Height          =   180
               Index           =   0
               Left            =   13155
               TabIndex        =   10
               Top             =   375
               Width           =   990
               _Version        =   786432
               _ExtentX        =   1746
               _ExtentY        =   317
               _StockProps     =   79
               Caption         =   "„„·Êﬂ…"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton ChCarType 
               Height          =   375
               Index           =   1
               Left            =   13395
               TabIndex        =   12
               Top             =   570
               Width           =   765
               _Version        =   786432
               _ExtentX        =   1349
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«Œ—Ï"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbCar 
               Height          =   315
               Left            =   7215
               TabIndex        =   11
               Top             =   375
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbCar2 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   675
               Width           =   6105
               _ExtentX        =   10769
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbSupplem2 
               Height          =   315
               Left            =   7215
               TabIndex        =   16
               Top             =   1005
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DBCboClientName 
               Height          =   315
               Left            =   7215
               TabIndex        =   14
               Top             =   675
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbSupplem 
               Height          =   315
               Left            =   120
               TabIndex        =   107
               Top             =   300
               Width           =   6105
               _ExtentX        =   10769
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   240
               Index           =   19
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   1005
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·„·Õﬁ"
               Height          =   240
               Index           =   18
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   1005
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Height          =   240
               Index           =   17
               Left            =   5475
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1005
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—ﬂ»…"
               Height          =   195
               Index           =   3
               Left            =   12180
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   375
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ê—œ"
               Height          =   270
               Index           =   11
               Left            =   12045
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   675
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„·Õﬁ"
               Height          =   240
               Index           =   10
               Left            =   12540
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   1005
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„—ﬂ»« "
               Height          =   285
               Index           =   9
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   120
               Width           =   1530
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   555
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   255
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„·Õﬁ"
               Height          =   195
               Index           =   7
               Left            =   5715
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   375
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—ﬂ»…"
               Height          =   270
               Index           =   6
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   675
               Width           =   1170
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   1410
            Left            =   0
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   3105
            Width           =   14325
            _cx             =   25268
            _cy             =   2487
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
            Begin VB.TextBox distance 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   330
               Left            =   11520
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   435
               Width           =   1560
            End
            Begin VB.TextBox delayCuz 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   435
               Width           =   5670
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   450
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   90
               Top             =   840
               Width           =   12990
            End
            Begin VB.TextBox TxtOrderNo 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   435
               Width           =   2670
            End
            Begin MSDataListLib.DataCombo DcbCity 
               Height          =   315
               Left            =   6960
               TabIndex        =   17
               Top             =   45
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbCity2 
               Height          =   315
               Left            =   120
               TabIndex        =   18
               Top             =   45
               Width           =   5670
               _ExtentX        =   10001
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«›…"
               Height          =   195
               Index           =   33
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   510
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "”»» «· √ŒÌ—"
               Height          =   225
               Index           =   25
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   495
               Width           =   900
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   195
               Index           =   15
               Left            =   13170
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   960
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÃÂ… «· Õ„Ì·"
               Height          =   225
               Index           =   14
               Left            =   13110
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   90
               Width           =   1050
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   465
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÃÂ… «· Ê—Ìœ"
               Height          =   225
               Index           =   13
               Left            =   5850
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   90
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «„— «· Õ„Ì· „‰ «·⁄„Ì·"
               Height          =   255
               Index           =   12
               Left            =   9690
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   480
               Width           =   1740
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   1125
            Left            =   0
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   5460
            Width           =   14325
            _cx             =   25268
            _cy             =   1984
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
            Begin VB.TextBox TxtTypGoods 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   13920
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   0
               Visible         =   0   'False
               Width           =   5970
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
               Height          =   1065
               Left            =   3180
               TabIndex        =   92
               Top             =   0
               Width           =   9630
               _cx             =   16986
               _cy             =   1879
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
               GridLines       =   3
               GridLinesFixed  =   2
               GridLineWidth   =   5
               Rows            =   1
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmOrderUpload.frx":1B57A
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
               Begin VB.PictureBox Picture2 
                  BorderStyle     =   0  'None
                  Height          =   1635
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   1635
                  ScaleWidth      =   2925
                  TabIndex        =   93
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   2925
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1125
                     Left            =   30
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   94
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2115
                  End
                  Begin VB.Label Label10 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   0
                     Width           =   2445
                  End
               End
            End
            Begin ImpulseButton.ISButton CmdDelete 
               Height          =   375
               Left            =   1710
               TabIndex        =   97
               Top             =   495
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
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
               ButtonImage     =   "FrmOrderUpload.frx":1B66D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «·ﬂ„Ì…"
               Height          =   495
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   120
               Width           =   975
            End
            Begin VB.Label LBLSUM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   495
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   555
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   225
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·»÷«⁄…"
               Height          =   345
               Index           =   16
               Left            =   12450
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   330
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   900
            Left            =   0
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   4560
            Width           =   14295
            _cx             =   25215
            _cy             =   1588
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
            Begin VB.CheckBox chkStop 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«›"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   555
               Width           =   735
            End
            Begin VB.ComboBox carStatus 
               Height          =   315
               ItemData        =   "FrmOrderUpload.frx":1BC07
               Left            =   120
               List            =   "FrmOrderUpload.frx":1BC09
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   120
               Width           =   3345
            End
            Begin VB.ComboBox orderStatus 
               Height          =   315
               ItemData        =   "FrmOrderUpload.frx":1BC0B
               Left            =   4680
               List            =   "FrmOrderUpload.frx":1BC0D
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   120
               Width           =   2025
            End
            Begin VB.TextBox delayHours 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   465
               Width           =   2160
            End
            Begin VB.TextBox txtCountOrders 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   480
               Width           =   2040
            End
            Begin MSComCtl2.DTPicker XPDTimeOrder 
               Height          =   315
               Left            =   11040
               TabIndex        =   101
               Top             =   120
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144375810
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker ETA 
               Height          =   315
               Left            =   11040
               TabIndex        =   110
               Top             =   480
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144375810
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker startDate 
               Height          =   315
               Left            =   7680
               TabIndex        =   114
               Top             =   120
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144375809
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker EDA 
               Height          =   315
               Left            =   7680
               TabIndex        =   115
               Top             =   480
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144375809
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·„⁄œÂ/«·”Ì«—…"
               Height          =   195
               Index           =   32
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·√„—"
               Height          =   195
               Index           =   31
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   120
               Width           =   780
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "”«⁄«  «· √ŒÌ—"
               Height          =   195
               Index           =   30
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   540
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄ ··Ê’Ê·"
               Height          =   195
               Index           =   28
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   480
               Width           =   1650
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ »œ«Ì… «·—Õ·…"
               Height          =   195
               Index           =   27
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   120
               Width           =   1410
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Êﬁ  «·„ Êﬁ⁄ ··Ê’Ê·"
               Height          =   315
               Index           =   26
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   480
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·—œÊœ"
               Height          =   195
               Index           =   21
               Left            =   6780
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   510
               Width           =   780
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  »œ«Ì… «·—Õ·…"
               Height          =   195
               Index           =   20
               Left            =   12750
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   120
               Width           =   1170
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic11 
            Height          =   510
            Left            =   0
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   2580
            Width           =   14325
            _cx             =   25268
            _cy             =   900
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
            Begin VB.ComboBox cmbTypeRep 
               Height          =   315
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   60
               Width           =   2325
            End
            Begin MSDataListLib.DataCombo DcbTripStatus 
               Height          =   315
               Left            =   7200
               TabIndex        =   105
               Top             =   120
               Width           =   5205
               _ExtentX        =   9181
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·—œ"
               Height          =   240
               Index           =   34
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   120
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·—œ"
               Height          =   315
               Index           =   22
               Left            =   12615
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   150
               Width           =   1440
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   750
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Width           =   14295
         _cx             =   25215
         _cy             =   1323
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
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   705
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   0
            Width           =   14505
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   450
               TabIndex        =   38
               Top             =   120
               Width           =   405
               _ExtentX        =   714
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
               ButtonImage     =   "FrmOrderUpload.frx":1BC0F
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   915
               TabIndex        =   39
               Top             =   120
               Width           =   405
               _ExtentX        =   714
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
               ButtonImage     =   "FrmOrderUpload.frx":1BFA9
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1515
               TabIndex        =   40
               Top             =   120
               Width           =   405
               _ExtentX        =   714
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
               ButtonImage     =   "FrmOrderUpload.frx":1C343
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   2040
               TabIndex        =   41
               Top             =   120
               Width           =   405
               _ExtentX        =   714
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
               ButtonImage     =   "FrmOrderUpload.frx":1C6DD
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " Õ«·… «·ÿ·»"
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
               Index           =   0
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   120
               Width           =   3240
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «„—  Õ„Ì·"
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
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   120
               Width           =   4080
            End
            Begin VB.Image Image1 
               Height          =   615
               Left            =   13200
               Picture         =   "FrmOrderUpload.frx":1CA77
               Stretch         =   -1  'True
               Top             =   0
               Width           =   735
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   630
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   765
         Width           =   14235
         _cx             =   25109
         _cy             =   1111
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
         Begin VB.TextBox TxtOrderStuts 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   390
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   390
            Left            =   10755
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   120
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   390
            Left            =   7770
            TabIndex        =   1
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   688
            _Version        =   393216
            Format          =   144375809
            CurrentDate     =   36526
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Bindings        =   "FrmOrderUpload.frx":1DE7C
            Height          =   315
            Left            =   4425
            TabIndex        =   2
            Top             =   120
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
         Begin MSDataListLib.DataCombo DBCboClientName1 
            Height          =   315
            Left            =   240
            TabIndex        =   88
            Top             =   120
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   360
            Index           =   24
            Left            =   3345
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   120
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   360
            Index           =   0
            Left            =   6450
            TabIndex        =   35
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Index           =   2
            Left            =   9765
            TabIndex        =   34
            Top             =   135
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   360
            Index           =   4
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   1275
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
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmOrderUpload.frx":1DE91
      Left            =   15480
      List            =   "FrmOrderUpload.frx":1DEA1
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   23
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
      RightToLeft     =   -1  'True
      TabIndex        =   22
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
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   25
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
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
      TabIndex        =   26
      Top             =   2280
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
            Picture         =   "FrmOrderUpload.frx":1DEBA
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1E254
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1E5EE
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1E988
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1ED22
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1F0BC
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1F456
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrderUpload.frx":1F9F0
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   27
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
      ButtonImage     =   "FrmOrderUpload.frx":1FD8A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
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
      ButtonImage     =   "FrmOrderUpload.frx":265EC
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
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
      ButtonImage     =   "FrmOrderUpload.frx":2CE4E
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
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmOrderUpload"
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
 Function GetGropID(Optional VehicleID As Double) As Double
 Dim sql As String
 Dim rs2 As ADODB.Recordset
 Set rs2 = New ADODB.Recordset
 sql = " SELECT     CarsTypeId, id"
 sql = sql & "   From dbo.TblCarsData"
 sql = sql & " Where (ID = " & VehicleID & ")"
 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
 GetGropID = IIf(IsNull(rs2("CarsTypeId").value), 0, rs2("CarsTypeId").value)
 Else
 GetGropID = 0
 End If
 End Function
Function GetGropID2(Optional VehicleID As Double) As Double
 Dim sql As String
 Dim rs2 As ADODB.Recordset
 Set rs2 = New ADODB.Recordset
 sql = " SELECT     ID, BrandID"
 sql = sql & "  From dbo.TblVendorCars"
 sql = sql & " Where (ID = " & VehicleID & ")"
 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
 GetGropID2 = IIf(IsNull(rs2("BrandID").value), 0, rs2("BrandID").value)
 Else
 GetGropID2 = 0
 End If
 End Function
Sub RetriveClinCounr(Optional VehicleID As Double, Optional Typ As Integer = 0)
If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim VehicleType As Double
If Typ = 0 Then
VehicleType = GetGropID(VehicleID)
Else
VehicleType = GetGropID2(VehicleID)
End If
sql = " SELECT     dbo.TblClientTransContrDet.Price, dbo.TblClientTransContrDet.Typed"
sql = sql & " FROM         dbo.TblClientTransContr LEFT OUTER JOIN"
sql = sql & "                      dbo.TblClientTransContrDet ON dbo.TblClientTransContr.ID = dbo.TblClientTransContrDet.ClintTransID"
sql = sql & " Where (dbo.TblClientTransContrDet.VehicleType = " & VehicleType & ") And (dbo.TblClientTransContr.LockedID = 0 or dbo.TblClientTransContr.LockedID is null)  And (dbo.TblClientTransContr.CompID = " & val(DBCboClientName1.BoundText) & ")"
sql = sql & " and dbo.TblClientTransContr.FromDate <=" & SQLDate(XPDtbTrans.value, True) & ""
sql = sql & " and dbo.TblClientTransContr.Todate >=" & SQLDate(XPDtbTrans.value, True) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
txtPrice.Text = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
Else
txtPrice.Text = 0
End If
End If
End Sub
Public Function saveGridData()
    Dim i As Integer
    Dim rss As ADODB.Recordset
    Set rss = New ADODB.Recordset
    Dim sql_str As String
    Dim StrSQL As String
    If Me.TxtModFlg.Text <> "R" Then
        If TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TravKItemDet1 Where MasterID = " & val(TxtSerial1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
          sql_str = "select * from TravKItemDet1 where 1=-1"
            rss.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    With Me.VSFlexGrid2
        For i = .FixedRows To .Rows - 1
            rss.AddNew
            rss("MasterID").value = val(TxtSerial1.Text)
            rss("ItemID").value = IIf((.TextMatrix(i, .ColIndex("KItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("KItemID"))))
            rss("Count").value = IIf((.TextMatrix(i, .ColIndex("Count"))) = "", Null, val((.TextMatrix(i, .ColIndex("Count")))))
            rss("UnitID").value = IIf((.TextMatrix(i, .ColIndex("KUnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("KUnitID"))))
            rss.update
        Next i
    End With
  End If
End Function
Sub fillItemsGrid()

    Dim i As Integer
    Dim rs_ItemsGrid As ADODB.Recordset
    Set rs_ItemsGrid = New ADODB.Recordset
    Dim StrSQL As String
        
    StrSQL = " SELECT     dbo.TravKItemDet1.MasterID, dbo.TravKItemDet1.[Count], dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TravKItemDet1.ItemID, "
    StrSQL = StrSQL & "                  dbo.TravKItemDet1.UnitID , dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL & "     FROM         dbo.TravKItemDet1 INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.TravKItemDet1.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where dbo.TravKItemDet1.MasterID = " & val(TxtSerial1.Text)
    'MasterID = " & val(TxtSerial1.Text)
    rs_ItemsGrid.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid2.Rows = 1
    
    If rs_ItemsGrid.RecordCount > 0 Then
        rs_ItemsGrid.MoveFirst
        With VSFlexGrid2
            .Rows = rs_ItemsGrid.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("KItemID")) = IIf(IsNull(rs_ItemsGrid("ItemID").value), 0, rs_ItemsGrid("ItemID").value)
                .TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(rs_ItemsGrid("Count").value), 0, rs_ItemsGrid("Count").value)
                .TextMatrix(i, .ColIndex("KUnitID")) = IIf(IsNull(rs_ItemsGrid("UnitID").value), 0, rs_ItemsGrid("UnitID").value)
               ' .TextMatrix(i, .ColIndex("nameE")) = IIf(IsNull(rs_ItemsGrid("namee").value), "", rs_ItemsGrid("namee").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("KItem")) = IIf(IsNull(rs_ItemsGrid("ItemName").value), "", rs_ItemsGrid("ItemName").value)
                    .TextMatrix(i, .ColIndex("KUnit")) = IIf(IsNull(rs_ItemsGrid("UnitName").value), "", rs_ItemsGrid("UnitName").value)
                Else
                    .TextMatrix(i, .ColIndex("KItem")) = IIf(IsNull(rs_ItemsGrid("ItemNamee").value), "", rs_ItemsGrid("ItemNamee").value)
                    .TextMatrix(i, .ColIndex("KUnit")) = IIf(IsNull(rs_ItemsGrid("UnitNamee").value), "", rs_ItemsGrid("UnitNamee").value)
                End If
                rs_ItemsGrid.MoveNext
            Next
        End With
    End If
    ReLineGrid
End Sub

Private Sub ChCarType_Click(Index As Integer)
If ChCarType(0).value = True Then
DcbSupplem.Enabled = True
DcbCar.Enabled = True
TxtSearchCode.Enabled = False
TxtSearchCode.Text = ""
DBCboClientName.Enabled = False
DcbCar2.Enabled = False
DcbSupplem2.Enabled = False
'TxtPrice.Enabled = False
txtPrice.Text = ""
TxtPartPrice.Enabled = False
TxtPartPrice.Text = ""
txtTotal.Enabled = False
txtTotal.Text = ""
Else
txtTotal.Enabled = True
TxtPartPrice.Enabled = True
txtPrice.Enabled = True
DcbSupplem2.Enabled = True
DcbSupplem2.BoundText = 0
DcbCar2.BoundText = 0
DcbCar2.Enabled = True
DBCboClientName.BoundText = 0
DBCboClientName.Enabled = True
TxtSearchCode.Enabled = True
DcbCar.BoundText = 0
DcbCar.Enabled = False
DcbSupplem.BoundText = 0
DcbSupplem.Enabled = False
End If

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
MySQL = " SELECT     dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate,TblOrderUpload.Price, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
MySQL = MySQL & "                      dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblOrderUpload.IDNo, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID,"
MySQL = MySQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods,"
MySQL = MySQL & "                      dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.PartPrice, dbo.TblOrderUpload.Price, dbo.TblOrderUpload.Total,"
MySQL = MySQL & "                      dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, dbo.TblOrderUpload.CityID2,"
MySQL = MySQL & "                      TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
MySQL = MySQL & "                      TblVendorCars_2.BoardNo AS BoardNo2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName, dbo.FixedAssets.namee AS SupplemNameE,"
MySQL = MySQL & "                      dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2,"
MySQL = MySQL & "                      TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2, dbo.TravKItemDet1.[Count], dbo.TravKItemDet1.ItemID,"
MySQL = MySQL & "                      dbo.TblItems.itemcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TravKItemDet1.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
MySQL = MySQL & " FROM         dbo.TblUnites RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblOrderUpload ON dbo.TravKItemDet1.MasterID = dbo.TblOrderUpload.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblOrderUpload.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblOrderUpload.ID =" & val(TxtSerial1.Text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaod.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderUplaod.rpt"
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
        Msg = "No Data"
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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
Private Sub RemoveGridRow()
    With Me.VSFlexGrid2
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub

Private Sub CmdDelete_Click()
If Me.TxtModFlg.Text <> "R" Then
RemoveGridRow
End If
End Sub
Private Sub DBCboClientName_Change()
DBCboClientName_Click (0)
End Sub
Private Sub DBCboClientName_Click(Area As Integer)
    Dim Fullcode As String
     Dim Dcombos As New ClsDataCombos
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 2
    TxtSearchCode.Text = Fullcode
     Dcombos.GetCarByVonder DcbCar2, val(DBCboClientName.BoundText)
End Sub
Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 100
        FrmCompanySearch.show vbModal
    
    End If
End Sub
Private Sub DBCboClientName1_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(DBCboClientName1.BoundText) <> 0 Then
If ChCarType(0).value = True Then
RetriveClinCounr val(val(DcbCar.BoundText)), 0
Else
RetriveClinCounr val(val(DcbCar2.BoundText)), 1
End If
End If
End If
End Sub

Private Sub DBCboClientName1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 101
        FrmCustemerSearch.show vbModal

    End If
End Sub

Private Sub DcbCar_Change()
DcbCar_Click (0)
End Sub

Private Sub DcbCar_Click(Area As Integer)
 Dim Dcombos As New ClsDataCombos
Dcombos.GetPartCar DcbSupplem, val(DcbCar.BoundText)
RetriveClinCounr val(val(DcbCar.BoundText)), 0
RetriveCarsInfo val(DcbCar.BoundText), 0
End Sub
Sub RetriveCarsInfo(Optional CarID As Double = 0, Optional Emp_id As String, Optional Typ As Integer = 0)
If Me.TxtModFlg <> "R" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblCarsData"
If Typ = 0 Then
sql = sql & "  Where id = " & CarID & ""
ElseIf Typ = 1 Then
sql = sql & " where Emp_id=" & Emp_id & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
If Typ <> 0 Then
DcbCar.BoundText = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
End If
DCEmployee.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
If Typ <> 0 Then
DcbCar.BoundText = 0
End If
End If
End If
End Sub
Private Sub DcbCar_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "OrderUpload"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbCar2_Change()
DcbCar2_Click (0)
End Sub
Private Sub DcbCar2_Click(Area As Integer)
Dim Dcombos As New ClsDataCombos
Dcombos.GetBartCarByVonder DcbSupplem2, val(DcbCar2.BoundText)
If Me.TxtModFlg.Text <> "R" Then
RetriveClinCounr val(val(DcbCar2.BoundText)), 1
Calc
End If
End Sub
Private Sub DcbCity_Change()
    GetTripInformations
End Sub
Private Sub DcbCity2_Change()
    GetTripInformations
End Sub

Private Sub DcbSupplem2_Change()
DcbSupplem2_Click (0)
End Sub

Private Sub DcbSupplem2_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
Calc
End If
End Sub

Private Sub DcEmployee_Change()
DcEmployee_Click (0)
End Sub

Private Sub DcEmployee_Click(Area As Integer)
Dim Nationality As String
Dim NumEkama As String
    If val(DCEmployee.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DCEmployee.BoundText, EmpCode
      Text6.Text = EmpCode
       If Me.TxtModFlg = "R" Then Exit Sub
        get_employee_information val(Me.DCEmployee.BoundText), , , , , , , , , Nationality, , , , , NumEkama
        TxtNationality.Text = Nationality
        TxtIDNo.Text = NumEkama
       ' RetriveCarsInfo val(DcEmployee.BoundText), 1
End Sub
Private Sub ChDrievType_Click(Index As Integer)
If ChDrievType(0).value = True Then
Text6.Enabled = True
DCEmployee.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.Text = ""
ElseIf ChDrievType(1).value = True Then
Text6.Enabled = False
DCEmployee.Enabled = False
TxtLeaderName.Enabled = True
DCEmployee.BoundText = 0
Text6.Text = ""
End If
End Sub
Sub LodR()
Dim str As String
  If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
   If SystemOptions.ShowDriverOnly = True Then
   str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
   End If
    fill_combo DCEmployee, str

End Sub
Function CheckTravel() As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblOrderUpload  where ID=" & val(TxtSerial1.Text) & " and IsTravel=1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckTravel = True
Else
CheckTravel = False
End If
End Function



Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Dim conection As String
    Dim My_SQL As String

    conection = "select * from TblOrderUpload  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    
   'load tblUsers -----------------------------------------------
    LodR
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.DcbBranch
    Dcombos.GetCitiesDistance Me.DcbCity, 0
    Dcombos.GetCitiesDistance Me.DcbCity2, 1
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName1
    Dcombos.GetEmployees Me.DcEmpSuper, True


    Dcombos.GetTripStatus Me.DcbTripStatus
    Dcombos.GetCars Me.DcbCar
   'Dcombos.GetCars Me.DcbCar2
   
    If SystemOptions.UserInterface = ArabicInterface Then
        With carStatus
            .Clear
            .AddItem "€Ì— „Õœœ"
            .AddItem "»«·ÿ—Ìﬁ"
            .AddItem "»«·„Êﬁ⁄"
            .AddItem "›«—€"
            .AddItem "»«·Ê—‘…"
        End With
        With cmbTypeRep
            .Clear
            .AddItem "œ«Œ·Ï"
            .AddItem "·Êﬂ«·"
            .AddItem "Œ«—ÃÌ"
        End With


        With orderStatus
            .Clear
            .AddItem "„› ÊÕ"
            .AddItem " „"
            .AddItem "„€·ﬁ"
            .AddItem " √ŒÌ—"
        End With
    Else
        With cmbTypeRep
            .Clear
            .AddItem "Internal"
            .AddItem "Local"
            .AddItem "External"
        End With
    
        With carStatus
            .Clear
            .AddItem "not defined"
            .AddItem "in the road"
            .AddItem "in the site"
            .AddItem "empty"
            .AddItem "in the workshop"
        End With

        With orderStatus
            .Clear
            .AddItem "Open"
            .AddItem "Done"
            .AddItem "Closed"
            .AddItem "Delayed"
        End With
    End If
    
    BtnLast_Click
 
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
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
     On Error GoTo ErrTrap
    Dim sql As String

    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.DcbBranch.BoundText)
    RsSavRec.Fields("EmpID").value = val(DCEmployee.BoundText)
    RsSavRec.Fields("IDNo").value = TxtIDNo.Text
    RsSavRec.Fields("LeaderName").value = TxtLeaderName.Text
    RsSavRec.Fields("Nationality").value = TxtNationality.Text
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("SupplemID").value = val(Me.DcbSupplem.BoundText)
    RsSavRec.Fields("SupplemID2").value = val(Me.DcbSupplem2.BoundText)
    RsSavRec.Fields("TripStatusID").value = val(Me.DcbTripStatus.BoundText)
    RsSavRec.Fields("CarID").value = val(Me.DcbCar.BoundText)
    RsSavRec.Fields("CarID2").value = val(Me.DcbCar2.BoundText)
    RsSavRec.Fields("CusID").value = val(Me.DBCboClientName.BoundText)
    
    RsSavRec.Fields("CustId1").value = val(Me.DBCboClientName1.BoundText)
    
    RsSavRec.Fields("CityID").value = val(Me.DcbCity.BoundText)
    RsSavRec.Fields("CityID2").value = val(Me.DcbCity2.BoundText)
    RsSavRec.Fields("TypGoods").value = TxtTypGoods.Text
    RsSavRec.Fields("OrderNo").value = txtOrderNo.Text
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("PartPrice").value = val(TxtPartPrice.Text)
    RsSavRec.Fields("Price").value = val(txtPrice.Text)
    RsSavRec.Fields("Total").value = val(txtTotal.Text)
    
    RsSavRec.Fields("CountOrders").value = val(txtCountOrders.Text)
    RsSavRec.Fields("TimeOrder").value = XPDTimeOrder.value
    
    RsSavRec.Fields("ETA").value = ETA.value
    RsSavRec.Fields("startDate").value = startDate.value
    RsSavRec.Fields("EDA").value = EDA.value
    RsSavRec.Fields("delayHours").value = val(delayHours.Text)
    'RsSavRec.Fields("supervisorName").value = supervisorName.Text
    RsSavRec.Fields("delayCuz").value = delayCuz.Text
    RsSavRec.Fields("orderStatus").value = IIf(val(Me.orderStatus.ListIndex) = -1, Null, Me.orderStatus.ListIndex)
    RsSavRec.Fields("carStatus").value = IIf(val(Me.carStatus.ListIndex) = -1, Null, Me.carStatus.ListIndex)
    RsSavRec.Fields("TypeRep").value = IIf(val(Me.cmbTypeRep.ListIndex) = -1, Null, Me.cmbTypeRep.ListIndex)
    
    RsSavRec.Fields("chkStop").value = chkStop.value
    RsSavRec.Fields("distance").value = val(distance.Text)
    RsSavRec.Fields("DcEmpSuper").value = val(Me.DcEmpSuper.BoundText)
    
    
   If ChDrievType(1).value = True Then
    RsSavRec.Fields("DrievType").value = 1
    Else
    RsSavRec.Fields("DrievType").value = 0
    End If
    If ChCarType(1).value = True Then
    RsSavRec.Fields("CarType").value = 1
    Else
    RsSavRec.Fields("CarType").value = 0
    End If
    RsSavRec.update
    saveGridData
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                FiLLTXT
                TxtModFlg = "R"
            If SystemOptions.UserInterface = ArabicInterface Then
             Else
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
            
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              '  Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
               ' Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
       
       RsSavRec.Resync adAffectCurrent
       
       
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()

   On Error GoTo ErrTrap
    'ChDrievType_Click (0)
    'ChCarType_Click (0)
    If Not IsNull(RsSavRec.Fields("DrievType").value) Then
        If val(RsSavRec.Fields("DrievType").value) = 1 Then
            ChDrievType(1).value = True
        Else
            ChDrievType(0).value = True
        End If
    Else
        ChDrievType(0).value = True
    End If
    
    If Not IsNull(RsSavRec.Fields("CarType").value) Then
        If val(RsSavRec.Fields("CarType").value) = 1 Then
            ChCarType(1).value = True
        Else
            ChCarType(0).value = True
        End If
    Else
        ChCarType(0).value = True
    End If
    ''///////////
        If Not IsNull(RsSavRec.Fields("OrderStuts").value) Then
        If val(RsSavRec.Fields("OrderStuts").value) = 1 Then
            Label1(0).Caption = "Õ«·… «·«„— „ﬁ›·"
        Else
           Label1(0).Caption = ""
        End If
    Else
        Label1(0).Caption = ""
    End If
    TxtOrderStuts.Text = IIf(IsNull(RsSavRec.Fields("OrderStuts").value), 0, RsSavRec.Fields("OrderStuts").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Me.DcbBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    Me.DBCboClientName1.BoundText = IIf(IsNull(RsSavRec.Fields("CustId1").value), "", RsSavRec.Fields("CustId1").value)
     
    Me.DCEmployee.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    TxtIDNo.Text = IIf(IsNull(RsSavRec.Fields("IDNo").value), "", RsSavRec.Fields("IDNo").value)
    TxtLeaderName.Text = IIf(IsNull(RsSavRec.Fields("LeaderName").value), "", RsSavRec.Fields("LeaderName").value)
    TxtNationality.Text = IIf(IsNull(RsSavRec.Fields("Nationality").value), "", RsSavRec.Fields("Nationality").value)
    Me.DcbCar.BoundText = IIf(IsNull(RsSavRec.Fields("CarID").value), "", RsSavRec.Fields("CarID").value)
    Me.DcbCar2.BoundText = IIf(IsNull(RsSavRec.Fields("CarID2").value), "", RsSavRec.Fields("CarID2").value)
    Me.DcbCity.BoundText = IIf(IsNull(RsSavRec.Fields("CityID").value), "", RsSavRec.Fields("CityID").value)
    Me.DcbCity2.BoundText = IIf(IsNull(RsSavRec.Fields("CityID2").value), "", RsSavRec.Fields("CityID2").value)
    TxtTypGoods.Text = IIf(IsNull(RsSavRec.Fields("TypGoods").value), "", RsSavRec.Fields("TypGoods").value)
    txtOrderNo.Text = IIf(IsNull(RsSavRec.Fields("OrderNo").value), "", RsSavRec.Fields("OrderNo").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbSupplem.BoundText = IIf(IsNull(RsSavRec.Fields("SupplemID").value), "", RsSavRec.Fields("SupplemID").value)
    Me.DcbSupplem2.BoundText = IIf(IsNull(RsSavRec.Fields("SupplemID2").value), "", RsSavRec.Fields("SupplemID2").value)
    Me.DcbTripStatus.BoundText = IIf(IsNull(RsSavRec.Fields("TripStatusID").value), "", RsSavRec.Fields("TripStatusID").value)
    
    TxtPartPrice.Text = IIf(IsNull(RsSavRec.Fields("PartPrice").value), "", RsSavRec.Fields("PartPrice").value)
    txtPrice.Text = IIf(IsNull(RsSavRec.Fields("Price").value), "", RsSavRec.Fields("Price").value)
    txtTotal.Text = IIf(IsNull(RsSavRec.Fields("Total").value), "", RsSavRec.Fields("Total").value)

    txtCountOrders.Text = IIf(IsNull(RsSavRec.Fields("CountOrders").value), "", RsSavRec.Fields("CountOrders").value)
    XPDTimeOrder.value = IIf(IsNull(RsSavRec.Fields("TimeOrder").value), Time, RsSavRec.Fields("TimeOrder").value)
    
    ETA.value = IIf(IsNull(RsSavRec.Fields("ETA").value), Time, RsSavRec.Fields("ETA").value)
    startDate.value = IIf(IsNull(RsSavRec.Fields("startDate").value), Date, RsSavRec.Fields("startDate").value)
    EDA.value = IIf(IsNull(RsSavRec.Fields("EDA").value), Date, RsSavRec.Fields("EDA").value)
    orderStatus.ListIndex = IIf(IsNull(RsSavRec.Fields("orderStatus").value), -1, RsSavRec.Fields("orderStatus").value)
    
    cmbTypeRep.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeRep").value), -1, RsSavRec.Fields("TypeRep").value)
    
    carStatus.ListIndex = IIf(IsNull(RsSavRec.Fields("carStatus").value), -1, RsSavRec.Fields("carStatus").value)
    delayHours.Text = IIf(IsNull(RsSavRec.Fields("delayHours").value), 0, RsSavRec.Fields("delayHours").value)
    'supervisorName.Text = IIf(IsNull(RsSavRec.Fields("supervisorName").value), "", RsSavRec.Fields("supervisorName").value)
    delayCuz.Text = IIf(IsNull(RsSavRec.Fields("delayCuz").value), "", RsSavRec.Fields("delayCuz").value)
    
    chkStop.value = IIf(IsNull(RsSavRec.Fields("chkStop").value), False, RsSavRec.Fields("chkStop").value)
    distance.Text = IIf(IsNull(RsSavRec.Fields("distance").value), 0, RsSavRec.Fields("distance").value)
    Me.DcEmpSuper.BoundText = IIf(IsNull(RsSavRec.Fields("DcEmpSuper").value), 0, RsSavRec.Fields("DcEmpSuper").value)
     
    fillItemsGrid
    
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
ErrTrap:
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()

   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    If val(Me.DcbBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
        Else
          MsgBox "Please select Branch"
        End If
        Exit Sub
    End If

 If val(Me.DBCboClientName1.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
        Else
          MsgBox "Please select Customer"
        End If
        Exit Sub
    End If
    If val(txtPrice) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Ì—ÃÏ «œŒ«· «·ﬁÌ„…"
        Else
          MsgBox "Please  Enter Value"
        End If
        Exit Sub
    
    End If
    If val(Me.DcEmpSuper.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Ì—ÃÏ «Œ Ì«— «·„—«ﬁ»"
        Else
          MsgBox "Please select Auditor"
        End If
        Exit Sub
    End If
    

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
 MsgBox "Erorr ... douring enter data", vbOKOnly + vbMsgBoxRight, App.title
End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblOrderUpload", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


Private Sub PartPrice_Change()
Calc
End Sub

Private Sub ISButton2_Click()
 On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
          
ShowAttachments TxtSerial1, "0411201802"
 
 

End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DCEmployee.BoundText = EmpID
    End If
End Sub

Private Sub txtPrice_Change()
Calc
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 2
        DBCboClientName.BoundText = CUSTID
    End If

End Sub
' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long, Optional ByVal NoteID As Double = 0)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
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

Private Sub btnDelete_Click()

              
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim i As Integer
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
     If val(TxtOrderStuts.Text) = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õ«·… «·«„— „ﬁ›· ·«Ì„ﬂ‰ Õ–›Â"
    Else
    MsgBox "You can not delete. This process is  locked"
    End If
    Exit Sub
    End If
    If CheckTravel() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› .Â–Â «·Õ—ﬂ… „— »ÿ… »«·—Õ·« "
    Else
    MsgBox "You can not delete. This process is linked to trips"
    End If
    Exit Sub
    End If
    Dim X As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
        
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            'StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
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
       ' Frm2.Enabled = True
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
      '  Grid.Enabled = True
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

      ' Frm2.Enabled = True
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
         '   Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
         '   Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
         '   Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
         If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    If val(TxtOrderStuts.Text) = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Õ«·… «·«„— „ﬁ›· ·«Ì„ﬂ‰  ⁄œÌ·Â"
    Else
    MsgBox "You can not edit. This process is  locked"
    End If
    Exit Sub
    End If
    If CheckTravel() = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· .Â–Â «·Õ—ﬂ… „— »ÿ… »«·—Õ·« "
    Else
    MsgBox "You can not edit. This process is linked to trips"
    End If
    Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
              VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
            VSFlexGrid2.Enabled = True


    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
           ' Msg = "⁄›Ê«" & Chr(13)
           ' Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
           ' Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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
   ' Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.DcbBranch.BoundText = Current_branch
   XPDtbTrans.value = Date
   ChDrievType(0).value = True
   ChCarType(0).value = True
   ChDrievType_Click (0)
   ChCarType_Click (0)
        VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = 2
            
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
        
           ' Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
           ' Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
           ' Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
           ' Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
           ' Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
           ' Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
Label1(2).Caption = "Order Loading"
Me.Caption = Label1(2).Caption
Label2(0).Caption = "Record No."
Label2(1).Caption = "Curr.Record"
    btnNew.Caption = "New"
    btnModify.Caption = "Edit"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    ISButton8.Caption = "Search"
    ISButton5.Caption = "Print"
    btnCancel.Caption = "Exit"
    CmdDelete.Caption = "Delete"

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    lbl(19).Caption = "Total"
    lbl(4).Caption = "ID"
    lbl(2).Caption = "Date"
    lbl(0).Caption = "Branch"
    lbl(24).Caption = "Customer"
    lbl(29).Caption = "Driver Name"
    ChDrievType(0).RightToLeft = False
    ChDrievType(1).RightToLeft = False
    ChDrievType(0).Caption = "Internal"
    ChDrievType(1).Caption = "External"
    lbl(1).Caption = "ID No."
    lbl(5).Caption = "Nationality"
    lbl(9).Caption = "Vehicle Data"
    ChCarType(0).RightToLeft = False
    ChCarType(1).RightToLeft = False
    ChCarType(0).Caption = "Owned"
    ChCarType(1).Caption = "Other"
    lbl(3).Caption = "Vehicle"
    lbl(6).Caption = "Vehicle"
    lbl(11).Caption = "Supplier"
    lbl(10).Caption = "Part"
    lbl(7).Caption = "Part"
    lbl(17).Caption = "Value"
    lbl(14).Caption = "From"
    lbl(13).Caption = "To"
    lbl(15).Caption = "Remarks"
    lbl(12).Caption = "Customer Order"
    lbl(16).Caption = "Goods"
    Label7.Caption = "Total Qty"
    lbl(8).Caption = "By"
    
    With Me.VSFlexGrid2
        .TextMatrix(0, .ColIndex("LineNo")) = "Serial"
        .TextMatrix(0, .ColIndex("KItem")) = "Item Name"
        .TextMatrix(0, .ColIndex("KUnit")) = "Unit Name"
        .TextMatrix(0, .ColIndex("Count")) = "Qty"
    End With
End Sub
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
 '  If val(Me.TxtSerial1.text) <> 0 Then
 '      print_report
 '  End If
ErrTrap:
End Sub


Sub Calc()
txtTotal.Text = val(txtPrice.Text) + val(TxtPartPrice.Text)
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblOrderUpload"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim StrItemID As String
    Dim StrUnitID As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sgl As String

    With VSFlexGrid2
        Select Case .ColKey(Col)
            Case "LineNo"
                '.TextMatrix(Row, .ColIndex("LineNo")) = setID_Line
            Case "KItem"
                StrItemID = .ComboData
                LngRow = .FindRow(StrItemID, .FixedRows, .ColIndex("LineNo"), False, True)
                .TextMatrix(Row, .ColIndex("KItemID")) = StrItemID
                StrItemID = .TextMatrix(Row, .ColIndex("KItemID"))

            Case "KUnit"
                StrUnitID = .ComboData
                LngRow = .FindRow(StrUnitID, .FixedRows, .ColIndex("KUnitID"), False, True)
                .TextMatrix(Row, .ColIndex("KUnitID")) = StrUnitID
                StrUnitID = .TextMatrix(Row, .ColIndex("KUnitID"))
        End Select

        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

    End With

    ReLineGrid

End Sub
Function ReLineGrid()
Dim i As Double
Dim SumQty As Double
Dim IntCounter As Double
LblSum = 0
    With Me.VSFlexGrid2

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("KItemID"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
               SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Count")))
                LblSum.Caption = SumQty
            End If

        Next i

    End With

End Function
Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With VSFlexGrid2
        Select Case .ColKey(Col)

            Case "Count"
                .ComboList = ""
            Case "LineNo"
                Cancel = True
                
        End Select
    End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String
    
    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "KItem"
                StrSQL = "select * from tblitems  where GroupID in ( "
                StrSQL = StrSQL & " SELECT     GroupID "
                StrSQL = StrSQL & " From dbo.Groups"
                StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"

                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid2.BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = VSFlexGrid2.BuildComboList(rs, "ItemNamee", "ItemID")
                End If
           
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "KUnit"
                StrSQL = "select * from TblUnites"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                         
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList1 = VSFlexGrid2.BuildComboList(rs, "UnitName", "UnitID")
                Else
                    StrComboList1 = VSFlexGrid2.BuildComboList(rs, "UnitNamee", "UnitID")
                End If
           
                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If
                
                .ComboList = StrComboList1
        End Select

    End With
End Sub
Private Sub delayHours_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, delayHours.Text, 1)
End Sub
Function GetTripInformations()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    
        Dim sql As String
        Dim rs As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TBLCitiesDistance"
        sql = sql & " Where (CityFromId = " & val(DcbCity.BoundText) & ") And (CitytoId=" & val(DcbCity2.BoundText) & ")"
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            distance.Text = IIf(IsNull(rs("Distance").value), 0, rs("Distance").value)
            'TxtKmPrice = IIf(IsNull(rs("Desil").value), 0, rs("Desil").value)
            'TXTTravelPrice = IIf(IsNull(rs("TravelPrice").value), 0, rs("TravelPrice").value)
            'TxtDriverPercentage = IIf(IsNull(rs("DriverPercentage").value), 0, rs("DriverPercentage").value)
            'txtDriverValue = IIf(IsNull(rs("DriverValue").value), 0, rs("DriverValue").value)
        Else
            'TxtDistance = 0
            'TxtKmPrice = 0
            'TXTTravelPrice = 0
            'TxtDriverPercentage = 0
            'txtDriverValue = 0
        End If
    End If
End Function
Private Sub DcEmpSuper_Change()
    If val(DcEmpSuper.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcEmpSuper.BoundText, EmpCode
    Text2.Text = EmpCode
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.Text, EmpID
        DcEmpSuper.BoundText = EmpID
    End If
End Sub
