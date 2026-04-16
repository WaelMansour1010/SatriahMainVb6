VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmAqarReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10515
   Icon            =   "FrmAqarReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _cx             =   18441
      _cy             =   13944
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
      Caption         =   " ﬁ«—Ì— «·⁄ﬁ«—| ﬁ«—Ì— «—»«Õ «·«⁄ﬁ«—|Tab&3"
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
      Flags(2)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   7530
         Left            =   11100
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   45
         Width           =   10365
         _cx             =   18283
         _cy             =   13282
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
         Begin VB.CommandButton Command1 
            Caption         =   "„”Õ"
            Height          =   495
            Left            =   2520
            TabIndex        =   69
            Top             =   5880
            Width           =   1125
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   5325
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   480
            Width           =   10395
            Begin VB.TextBox TxtAmount 
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
               Left            =   3720
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   375
               Left            =   240
               TabIndex        =   72
               Top             =   2400
               Width           =   3375
               Begin XtremeSuiteControls.RadioButton opt2 
                  Height          =   255
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   73
                  Top             =   0
                  Width           =   495
                  _Version        =   786432
                  _ExtentX        =   873
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "<"
                  ForeColor       =   16711680
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton opt2 
                  Height          =   255
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   74
                  Top             =   0
                  Width           =   495
                  _Version        =   786432
                  _ExtentX        =   873
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   ">"
                  ForeColor       =   16711680
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton opt2 
                  Height          =   255
                  Index           =   2
                  Left            =   1440
                  TabIndex        =   75
                  Top             =   0
                  Width           =   495
                  _Version        =   786432
                  _ExtentX        =   873
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "="
                  ForeColor       =   16711680
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton opt2 
                  Height          =   255
                  Index           =   3
                  Left            =   720
                  TabIndex        =   76
                  Top             =   0
                  Width           =   495
                  _Version        =   786432
                  _ExtentX        =   873
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "<="
                  ForeColor       =   16711680
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton opt2 
                  Height          =   255
                  Index           =   4
                  Left            =   0
                  TabIndex        =   77
                  Top             =   0
                  Width           =   495
                  _Version        =   786432
                  _ExtentX        =   873
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   ">="
                  ForeColor       =   16711680
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  UseVisualStyle  =   -1  'True
               End
            End
            Begin VB.Frame Frame2 
               Height          =   5175
               Left            =   6720
               TabIndex        =   55
               Top             =   120
               Width           =   3615
               Begin VB.Image Image2 
                  Height          =   4110
                  Left            =   120
                  Picture         =   "FrmAqarReport.frx":038A
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   3540
               End
               Begin VB.Label Label3 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”« —Ì…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   27.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   3855
                  Left            =   360
                  TabIndex        =   56
                  Top             =   4200
                  Width           =   2895
               End
            End
            Begin VB.TextBox Text7 
               Height          =   285
               Left            =   6360
               TabIndex        =   54
               Top             =   5760
               Width           =   855
            End
            Begin VB.TextBox txtCodeOwner2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               TabIndex        =   53
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   5400
               TabIndex        =   52
               Top             =   5760
               Width           =   855
            End
            Begin MSDataListLib.DataCombo DcbBranch2 
               Height          =   315
               Left            =   240
               TabIndex        =   57
               Top             =   240
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcsupplier2 
               Height          =   315
               Left            =   240
               TabIndex        =   58
               Top             =   1320
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcbAqarType2 
               Height          =   315
               Left            =   240
               TabIndex        =   59
               Top             =   600
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcaqartypeid2 
               Height          =   315
               Left            =   240
               TabIndex        =   60
               Top             =   960
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcbCityId22 
               Height          =   315
               Left            =   240
               TabIndex        =   65
               Top             =   2040
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcmCityID2 
               Height          =   315
               Left            =   240
               TabIndex        =   66
               Top             =   1680
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " „»·€ «·—»Õ"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   18
               Left            =   5505
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   2400
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·ÕÏ „⁄Ì‰"
               Height          =   195
               Index           =   17
               Left            =   5400
               TabIndex        =   68
               Top             =   1680
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·„œÌ‰… „⁄Ì‰…"
               Height          =   195
               Index           =   16
               Left            =   5400
               TabIndex        =   67
               Top             =   2040
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·⁄ﬁ«—"
               Height          =   195
               Index           =   25
               Left            =   5340
               TabIndex        =   64
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·„«·ﬂ „Õœœ"
               Height          =   195
               Index           =   15
               Left            =   5400
               TabIndex        =   63
               Top             =   1320
               Width           =   1065
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ«·⁄ﬁ«— „⁄Ì‰"
               Height          =   195
               Index           =   13
               Left            =   5400
               TabIndex        =   62
               Top             =   600
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
               Height          =   195
               Index           =   12
               Left            =   5400
               TabIndex        =   61
               Top             =   240
               Width           =   1020
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   1320
            TabIndex        =   70
            Top             =   5880
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   71
            Top             =   5880
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   " ﬁ«—Ì— «—»Õ «·⁄ﬁ«—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   10335
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7530
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   10365
         _cx             =   18283
         _cy             =   13282
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
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Left            =   2760
            TabIndex        =   46
            Top             =   6990
            Width           =   1125
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Height          =   6495
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   10395
            Begin VB.CheckBox ChkNotAccredit 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄ﬁÊœ «·€Ì— «·„ÊÀﬁ… ›ﬁÿ"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   870
               Width           =   1905
            End
            Begin VB.ComboBox ServersName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   7020
               Visible         =   0   'False
               Width           =   3345
            End
            Begin VB.ComboBox dbname 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   6000
               Width           =   3345
            End
            Begin VB.CheckBox ChkAccredit 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄ﬁÊœ «·„ÊÀﬁ… ›ﬁÿ"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   600
               Width           =   1905
            End
            Begin VB.TextBox txtRoomCount 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               TabIndex        =   20
               Top             =   5520
               Width           =   4935
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   7560
               TabIndex        =   19
               Top             =   6030
               Width           =   855
            End
            Begin VB.TextBox txtCodeSalesRep 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               TabIndex        =   18
               Top             =   3000
               Width           =   855
            End
            Begin VB.TextBox txtCodeOwner 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               TabIndex        =   17
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txtCodeCustomer 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               TabIndex        =   16
               Top             =   2640
               Width           =   855
            End
            Begin VB.TextBox txtCodeBranch 
               Height          =   285
               Left            =   6990
               TabIndex        =   15
               Top             =   6030
               Width           =   855
            End
            Begin VB.Frame Frame3 
               Height          =   4425
               Left            =   6720
               TabIndex        =   13
               Top             =   120
               Width           =   3615
               Begin VB.Label lblCompanyname 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”« —Ì…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   27.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   5295
                  Left            =   120
                  TabIndex        =   14
                  Top             =   2520
                  Width           =   2895
               End
               Begin VB.Image Image1 
                  Height          =   2310
                  Left            =   120
                  Picture         =   "FrmAqarReport.frx":10A48
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   3300
               End
            End
            Begin VB.TextBox txtStreet 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               TabIndex        =   12
               Top             =   5160
               Width           =   4935
            End
            Begin VB.Frame XPPnlTime 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· «—ÌŒ "
               Height          =   1305
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   4560
               Width           =   3615
               Begin MSComCtl2.DTPicker XPDtbFrom 
                  Height          =   345
                  Left            =   1560
                  TabIndex        =   5
                  Top             =   360
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   609
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   144179201
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker XPDtpTo 
                  Height          =   345
                  Left            =   1560
                  TabIndex        =   6
                  Top             =   840
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   609
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   144179201
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal XPDtbFromH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   7
                  Top             =   360
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal XPDtpToH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   8
                  Top             =   840
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   556
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   285
                  Index           =   2
                  Left            =   3000
                  TabIndex        =   10
                  Top             =   360
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   285
                  Index           =   0
                  Left            =   3000
                  TabIndex        =   9
                  Top             =   840
                  Width           =   465
               End
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   0
               Left            =   5310
               TabIndex        =   11
               Top             =   240
               Width           =   1365
               _Version        =   786432
               _ExtentX        =   2408
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ«—Ì— «·⁄ﬁ«—« "
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch 
               Height          =   315
               Left            =   240
               TabIndex        =   21
               Top             =   1200
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcsupplier 
               Height          =   315
               Left            =   240
               TabIndex        =   22
               Top             =   2280
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcbAqarType 
               Height          =   315
               Left            =   240
               TabIndex        =   23
               Top             =   1560
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dbcClient 
               Height          =   315
               Left            =   240
               TabIndex        =   24
               Top             =   2640
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcbSalesSpec 
               Height          =   315
               Left            =   240
               TabIndex        =   25
               Top             =   3000
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dbcAqarStatus 
               Height          =   315
               Left            =   240
               TabIndex        =   26
               Top             =   3360
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCAkarUnit 
               Height          =   315
               Left            =   240
               TabIndex        =   27
               Top             =   4440
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcbCityId2 
               Height          =   315
               Left            =   240
               TabIndex        =   28
               Top             =   4080
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcmCityID 
               Height          =   315
               Left            =   240
               TabIndex        =   29
               Top             =   3720
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   1
               Left            =   2760
               TabIndex        =   30
               Top             =   240
               Width           =   1515
               _Version        =   786432
               _ExtentX        =   2672
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„’—Ê›«  «Œ—Ï"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitNo 
               Height          =   315
               Left            =   240
               TabIndex        =   31
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   4800
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcaqartypeid 
               Height          =   315
               Left            =   240
               TabIndex        =   32
               Top             =   1920
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   2
               Left            =   -570
               TabIndex        =   80
               Top             =   240
               Width           =   3255
               _Version        =   786432
               _ExtentX        =   5741
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ«—Ì— «·⁄ﬁ«—«  ÿ»ﬁ« ··⁄ﬁ«—"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   4
               Left            =   2910
               TabIndex        =   81
               Top             =   840
               Width           =   1395
               _Version        =   786432
               _ExtentX        =   2461
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«Ã„«·Ì «· ’›Ì…"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   5
               Left            =   4440
               TabIndex        =   82
               Top             =   840
               Width           =   2265
               _Version        =   786432
               _ExtentX        =   3995
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ«—Ì— «·⁄ﬁ«—«  ÿ»ﬁ« ··›—⁄"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁ«⁄œÂ «·»Ì«‰« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   159
               Left            =   5250
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   6060
               Width           =   1305
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
               Height          =   195
               Index           =   0
               Left            =   5400
               TabIndex        =   45
               Top             =   1200
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ«·⁄ﬁ«— „⁄Ì‰"
               Height          =   195
               Index           =   1
               Left            =   5400
               TabIndex        =   44
               Top             =   1560
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·„«·ﬂ „Õœœ"
               Height          =   195
               Index           =   2
               Left            =   5400
               TabIndex        =   43
               Top             =   2280
               Width           =   1065
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
               Height          =   195
               Index           =   3
               Left            =   5400
               TabIndex        =   42
               Top             =   2640
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·„‰œÊ» „Õœœ"
               Height          =   195
               Index           =   4
               Left            =   5400
               TabIndex        =   41
               Top             =   3000
               Width           =   1185
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·Õ«·… «·ÊÕœ…"
               Height          =   195
               Index           =   6
               Left            =   5340
               TabIndex        =   40
               Top             =   3360
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·ÕÏ „⁄Ì‰"
               Height          =   195
               Index           =   7
               Left            =   5400
               TabIndex        =   39
               Top             =   3720
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·„œÌ‰… „⁄Ì‰…"
               Height          =   195
               Index           =   8
               Left            =   5400
               TabIndex        =   38
               Top             =   4080
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·ÊÕœ…"
               Height          =   195
               Index           =   9
               Left            =   5400
               TabIndex        =   37
               Top             =   4440
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·‘«—⁄ „⁄Ì‰"
               Height          =   195
               Index           =   10
               Left            =   5400
               TabIndex        =   36
               Top             =   5160
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·⁄œœ «·€—›"
               Height          =   195
               Index           =   11
               Left            =   5400
               TabIndex        =   35
               Top             =   5520
               Width           =   1065
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÊÕœ…"
               Height          =   195
               Index           =   14
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   4800
               Width           =   1140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·⁄ﬁ«—"
               Height          =   195
               Index           =   5
               Left            =   5340
               TabIndex        =   33
               Top             =   1920
               Width           =   1080
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   1560
            TabIndex        =   47
            Top             =   6990
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   360
            TabIndex        =   48
            Top             =   6990
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   360
            Picture         =   "FrmAqarReport.frx":21106
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "‘«‘…  ﬁ«—Ì— «·⁄ﬁ«—« /«·„’—Ê›«  «·«Œ—Ï"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   840
            TabIndex        =   2
            Top             =   120
            Width           =   11295
         End
      End
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   600
      Picture         =   "FrmAqarReport.frx":24D6E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "FrmAqarReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mServer As String
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub btnClear_Click()
clear_all Me
XPDtbFrom.value = ""
XPDtpTo.value = ""
End Sub
Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
 If Rd(1).value = True Then
    GetDataExpen
 ElseIf Rd(0).value = True Then
    GetData
ElseIf Rd(2).value = True Or Rd(5).value = True Then
    GetData3
ElseIf Rd(4).value = True Then
    GetData4
 End If
        Case 1
      GetData2
        Case 2
            Unload Me
            Case 3
'print_report
    End Select
End Sub

Private Sub Command1_Click()
clear_all Me
End Sub


Private Sub dbname_Change()

ServersName.ListIndex = dbname.ListIndex
mServer = ServersName.Text & ".dbo."
    
End Sub

Private Sub dbname_Click()

    ServersName.ListIndex = dbname.ListIndex

End Sub

Private Sub dbname_Validate(Cancel As Boolean)
ServersName.ListIndex = dbname.ListIndex


End Sub


Private Sub DCAkarUnit_Change()
DCAkarUnit_Click (0)
End Sub

Private Sub DCAkarUnit_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
Set Dcombos = New ClsDataCombos
If val(dcbAqarType.BoundText) > 0 Then
idd = val(dcbAqarType.BoundText)
idd1 = val(DCAkarUnit.BoundText)
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
End If
End Sub

Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
DCAkarUnit_Change
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub
Private Sub Form_Load()
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Rd(0).value = True
    XPPnlTime.Visible = True
    XPDtbFrom.value = Date
    XPDtpTo.value = Date
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetIqar dcbAqarType
    Dcombos.GetCountriesGovernCities dcmCityID
    Dcombos.getCountriesGovernments dcbCityId2
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Dcombos.GetIqar dcbAqarType2
    Dcombos.GetCountriesGovernCities dcmCityID2
    Dcombos.getCountriesGovernments dcbCityId22
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier2
    Dcombos.getAkarUnit Me.DCAkarUnit
    Dcombos.GetSalesRepData Me.dcbSalesSpec
    Dcombos.getAkarType Me.dcaqartypeid
    Dcombos.getAkarType Me.dcaqartypeid2
    Dcombos.GetCustomersSuppliers 56, Me.dbcClient
    Dcombos.GetBranches DcbBranch2
    Dcombos.GetBranches DcbBranch
    Dcombos.GetRentStatus dbcAqarStatus
    Set cSearch = New clsDCboSearch
    My_SQL = "TblContract"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Resize_Form Me
    
        
On Error Resume Next
    
Dim a As Variant
Dim VarSet As Variant
  Open App.path & "\DB.txt" For Input As #1
    dbname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
              '  DbNamePath.AddItem (VarSet(0))
                dbname.AddItem (VarSet(1))
                ServersName.AddItem (VarSet(0))
                            
            End If
        End If
    
    Loop

    Close #1


End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetDataExpen()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = "SELECT     dbo.TblOtheExpensAqar.ID, dbo.TblOtheExpensAqar.RecordDateH, dbo.TblOtheExpensAqar.RecordDate, dbo.TblOtheExpensAqar.BranchID, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblOtheExpensAqar.Valuee, dbo.TblOtheExpensAqar.AqarID, dbo.TblAqar.aqarname,"
StrSQL = StrSQL & "                      dbo.TblAqar.aqarNo, dbo.TblOtheExpensAqar.UnitTypID, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblOtheExpensAqar.UnitID, dbo.TblAqarDetai.unitno,"
StrSQL = StrSQL & "                      dbo.TblOtheExpensAqar.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Cus_Phone,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.HomeTel, dbo.TblOtheExpensAqar.Mobile,"
StrSQL = StrSQL & "                      dbo.TblOtheExpensAqar.TypID, dbo.TblOtheExpensAqar.BillNo, dbo.TblOtheExpensAqar.AccountNo, dbo.TblOtheExpensAqar.AccountBank,"
StrSQL = StrSQL & "                      dbo.TblOtheExpensAqar.Remarks, dbo.TblOtheExpensAqar.PayedDateH, dbo.TblOtheExpensAqar.PayedDate, dbo.TblOtheExpensAqar.FromDateH,"
StrSQL = StrSQL & "                      dbo.TblOtheExpensAqar.FromDate, dbo.TblOtheExpensAqar.ToDateH, dbo.TblOtheExpensAqar.ToDate, dbo.TblOtheExpensAqar.EmpID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblAqar.ownerid, TblCustemers_1.CusName AS OwnerName, TblCustemers_1.CusNamee AS OwnerNameE,"
StrSQL = StrSQL & "                      TblCustemers_1.Fullcode AS OwFullcode, dbo.TblAqar.CountryID, dbo.TblAqar.cityid, dbo.TblAqar.streetname, dbo.TblAqarDetai.roomscount , dbo.TblAqarDetai.Status"
StrSQL = StrSQL & " FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON TblCustemers_1.CusID = dbo.TblAqar.ownerid RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblOtheExpensAqar LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblOtheExpensAqar.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblOtheExpensAqar.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblOtheExpensAqar.UnitID = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblOtheExpensAqar.UnitTypID = dbo.TblAkarUnit.id ON dbo.TblAqar.Aqarid = dbo.TblOtheExpensAqar.AqarID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblOtheExpensAqar.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "   Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
If Me.DcbUnitNo.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqarDetai.ID = " & val(Me.DcbUnitNo.BoundText)
End If
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblOtheExpensAqar.BranchID = " & val(Me.DcbBranch.BoundText)
End If
If Me.dcbAqarType.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.AqarID = " & val(Me.dcbAqarType.BoundText)
End If
If Me.dcsupplier.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.ownerid = " & val(dcsupplier.BoundText)
End If
If Me.dbcClient.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblOtheExpensAqar.CusID = " & val(dbcClient.BoundText)
End If
If Me.dcbCityId2.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CountryID = " & val(dcbCityId2.BoundText)
End If
If Me.dcmCityID.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CityID = " & val(dcmCityID.BoundText)
End If
If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblOtheExpensAqar.UnitTypID = " & val(DCAkarUnit.BoundText)
End If
If Me.dbcAqarStatus.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqarDetai.Status = " & val(dbcAqarStatus.BoundText)
End If
If Me.TxtStreet.Text <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqar.streetName like '%" & (TxtStreet.Text) & "%'"
End If
If Me.txtRoomCount.Text <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqarDetai.roomscount = " & val(txtRoomCount.Text)
End If

If Not IsNull(XPDtbFrom.value) Then
StrWhere = StrWhere & " AND  dbo.TblOtheExpensAqar.PayedDate >= " & SQLDate(XPDtbFrom.value, True) & ""
End If
If Not IsNull(XPDtpTo.value) Then
StrWhere = StrWhere & " AND  dbo.TblOtheExpensAqar.PayedDate <= " & SQLDate(XPDtpTo.value, True) & ""
End If
    StrSQL = StrSQL & StrWhere
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL, 1

            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
    End If
    End If
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = " SELECT     dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.RentValue, dbo.TblAqar.ownerid, "
StrSQL = StrSQL & "                      TblCustemers_1.CusID, TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname,"
StrSQL = StrSQL & "                      TblCustemers_2.CusName AS ownername, TblCustemers_2.CusNamee AS ownernamee, dbo.TblAqar.CountryID, dbo.TblAqar.cityid, dbo.TblAqar.streetname,"
StrSQL = StrSQL & "                      dbo.TblAkarUnit.name AS UnitName, dbo.TblAkarUnit.namee AS UnitNamee, dbo.TblRentStatus.name AS RentStatusName,"
StrSQL = StrSQL & "                      dbo.TblRentStatus.namee AS RentStatusNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Services,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.Water, dbo.TblAqarDetai.electric, dbo.GetLastDateUonit(dbo.TblAqarDetai.Id) AS MaxDate, dbo.TblAqarDetai.Id,"
StrSQL = StrSQL & "                      dbo.GetLastDateUonitRent(dbo.TblAqarDetai.Id) AS lastDate, dbo.TblRentStatus.id AS IdRent, dbo.TblAqarDetai.Comm, dbo.TblAqarDetai.InsuranceValue,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.customerid, dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.meterPrice,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.haveFurniture, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ContID, dbo.TblContract.PeriodsID, dbo.TblContract.Periods,"
StrSQL = StrSQL & "                      dbo.TblAqar.aqartypeid, dbo.tblAkarType.name AS AqarType, dbo.tblAkarType.namee AS AqarTypeE"
StrSQL = StrSQL & " FROM         dbo.TblAqarDetai INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblContract ON dbo.TblAqarDetai.ContID = dbo.TblContract.ContNo LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRentStatus ON dbo.TblAqarDetai.Status = dbo.TblRentStatus.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblAqarDetai.customerid = TblCustemers_1.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_2 ON dbo.TblAqar.ownerid = TblCustemers_2.CusID"
StrSQL = StrSQL & "   Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
    
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.BranchId = " & val(Me.DcbBranch.BoundText)
End If
If Me.dcaqartypeid.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid  = " & val(Me.dcaqartypeid.BoundText)
End If

If ChkAccredit.value = vbChecked Then
    StrWhere = StrWhere & " AND IsNull(TblContract.Accredit,0) = 1 "
End If

If ChkNotAccredit.value = vbChecked Then
    StrWhere = StrWhere & " AND IsNull(TblContract.Accredit,0) = 0 "
End If


If Not IsNull(XPDtbFrom.value) Then
    StrWhere = StrWhere & " AND  dbo.TblContract.ContDate >= " & SQLDate(XPDtbFrom.value, True) & ""
End If
If Not IsNull(XPDtpTo.value) Then
    StrWhere = StrWhere & " AND  dbo.TblContract.ContDate <= " & SQLDate(XPDtpTo.value, True) & ""
End If


If Me.DcbUnitNo.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqarDetai.ID = " & val(Me.DcbUnitNo.BoundText)
End If

If Me.dcbAqarType.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqarDetai.AqarID = " & val(Me.dcbAqarType.BoundText)
End If


If Me.dcsupplier.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.ownerid = " & val(dcsupplier.BoundText)
'gr = 2
End If

If Me.dbcClient.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqarDetai.customerid = " & val(dbcClient.BoundText)
'gr = 2
End If

'If Me.dcbSalesSpec.BoundText <> "" Then
'StrWhere = StrWhere & " AND cusid = " & val(dcbSalesSpec.BoundText)
'gr = 2
'End If


If Me.dcbCityId2.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CountryID = " & val(dcbCityId2.BoundText)
'gr = 2
End If

If Me.dcmCityID.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CityID = " & val(dcmCityID.BoundText)
'gr = 2
End If


If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqarDetai.unittype = " & val(DCAkarUnit.BoundText)
'gr = 2
End If


If Me.dbcAqarStatus.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqarDetai.Status = " & val(dbcAqarStatus.BoundText)
'gr = 2
End If


If Me.TxtStreet.Text <> "" Then
StrWhere = StrWhere & " AND  streetName like '%" & (TxtStreet.Text) & "%'"
'gr = 2
End If

If Me.txtRoomCount.Text <> "" Then
StrWhere = StrWhere & " AND  roomscount = " & val(txtRoomCount.Text)
'gr = 2
End If

If Me.XPDtbFrom.value = 1 Then
StrWhere = StrWhere & " AND  roomscount = " & val(txtRoomCount.Text)
'gr = 2
End If


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL

    End If

End Sub
Public Sub GetData2()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = " SELECT     dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.CountryID, dbo.TblAqar.cityid, dbo.TblAqar.streetname, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                       dbo.TblBranchesData.branch_namee, dbo.TblAqar.aqartypeid, dbo.TblAqar.ContValue, dbo.TblAqar.ownerid, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                       dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblAqar.PaymentNo, dbo.GetCountUnitMaint(dbo.TblAqar.Aqarid) AS UnitMaint,"
StrSQL = StrSQL & "                       dbo.GetCountUnitEmpty(dbo.TblAqar.Aqarid) AS UnitEmpty ,dbo.GetCountAllUnit(dbo.TblAqar.Aqarid) AS AllUnit, dbo.GetRentValueUnit(dbo.TblAqar.Aqarid) AS TotalRent"
StrSQL = StrSQL & "  FROM         dbo.TblAqar LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  Where (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
    
If Me.DcbBranch2.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.BranchId = " & val(Me.DcbBranch2.BoundText)
End If
If Me.dcaqartypeid2.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid  = " & val(Me.dcaqartypeid2.BoundText)
End If
If Me.dcbAqarType2.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType2.BoundText)
End If
If Me.dcsupplier2.BoundText <> "" Then
StrWhere = StrWhere & " AND ownerid = " & val(dcsupplier2.BoundText)
End If

If Me.dcbCityId22.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CountryID = " & val(dcbCityId22.BoundText)
End If
If Me.dcmCityID2.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CityID = " & val(dcmCityID2.BoundText)
End If
''//////////
If val(txtAmount.Text) <> 0 Then

If opt2(1).value = True Then
StrWhere = StrWhere & " AND (ISNULL(dbo.GetRentValueUnit(dbo.TblAqar.Aqarid), 0) - ISNULL(dbo.TblAqar.ContValue, 0) / CASE WHEN isnull(dbo.TblAqar.PaymentNo, 1) "
StrWhere = StrWhere & "                         = 0 THEN 1 ELSE dbo.TblAqar.PaymentNo END )  > " & val(txtAmount.Text)
ElseIf opt2(0).value = True Then
StrWhere = StrWhere & " AND (ISNULL(dbo.GetRentValueUnit(dbo.TblAqar.Aqarid), 0) - ISNULL(dbo.TblAqar.ContValue, 0) / CASE WHEN isnull(dbo.TblAqar.PaymentNo, 1) "
StrWhere = StrWhere & "                         = 0 THEN 1 ELSE dbo.TblAqar.PaymentNo END )  < " & val(txtAmount.Text)
ElseIf opt2(3).value = True Then
StrWhere = StrWhere & " AND (ISNULL(dbo.GetRentValueUnit(dbo.TblAqar.Aqarid), 0) - ISNULL(dbo.TblAqar.ContValue, 0) / CASE WHEN isnull(dbo.TblAqar.PaymentNo, 1) "
StrWhere = StrWhere & "                         = 0 THEN 1 ELSE dbo.TblAqar.PaymentNo END )  <= " & val(txtAmount.Text)
ElseIf opt2(4).value = True Then
StrWhere = StrWhere & " AND (ISNULL(dbo.GetRentValueUnit(dbo.TblAqar.Aqarid), 0) - ISNULL(dbo.TblAqar.ContValue, 0) / CASE WHEN isnull(dbo.TblAqar.PaymentNo, 1) "
StrWhere = StrWhere & "                         = 0 THEN 1 ELSE dbo.TblAqar.PaymentNo END )  >= " & val(txtAmount.Text)
ElseIf opt2(2).value = True Then
StrWhere = StrWhere & "  AND (ISNULL(dbo.GetRentValueUnit(dbo.TblAqar.Aqarid), 0) - ISNULL(dbo.TblAqar.ContValue, 0) / CASE WHEN isnull(dbo.TblAqar.PaymentNo, 1) "
StrWhere = StrWhere & "                         = 0 THEN 1 ELSE dbo.TblAqar.PaymentNo END )  = " & val(txtAmount.Text)
End If
End If

    StrSQL = StrSQL & StrWhere
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then


        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report2 StrSQL

    End If

End Sub


Public Sub GetData3()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = " SELECT     dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.ACCount,TblCustemers_1.Cus_mobile, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.RentValue, dbo.TblAqar.ownerid, "
StrSQL = StrSQL & "                      TblCustemers_1.CusID, TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarname,"
StrSQL = StrSQL & "                      TblCustemers_2.CusName AS ownername, TblCustemers_2.CusNamee AS ownernamee, dbo.TblAqar.CountryID, dbo.TblAqar.cityid, dbo.TblAqar.streetname,"
StrSQL = StrSQL & "                      dbo.TblAkarUnit.name AS UnitName, dbo.TblAkarUnit.namee AS UnitNamee, dbo.TblRentStatus.name AS RentStatusName,"
StrSQL = StrSQL & "                      dbo.TblRentStatus.namee AS RentStatusNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Services,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.Water, dbo.TblAqarDetai.electric, dbo.GetLastDateUonit(dbo.TblAqarDetai.Id) AS MaxDate, dbo.TblAqarDetai.Id,"
StrSQL = StrSQL & "                      dbo.GetLastDateUonitRent(dbo.TblAqarDetai.Id) AS lastDate, dbo.TblRentStatus.id AS IdRent, dbo.TblAqarDetai.Comm, dbo.TblAqarDetai.InsuranceValue,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.customerid, dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.meterPrice,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.haveFurniture, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ContID, dbo.TblContract.PeriodsID, dbo.TblContract.Periods,"
StrSQL = StrSQL & "                      dbo.TblAqar.aqartypeid, dbo.tblAkarType.name AS AqarType, dbo.tblAkarType.namee AS AqarTypeE"
StrSQL = StrSQL & " FROM         dbo.TblAqarDetai INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblContract ON dbo.TblAqarDetai.ContID = dbo.TblContract.ContNo LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRentStatus ON dbo.TblAqarDetai.Status = dbo.TblRentStatus.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblAqarDetai.customerid = TblCustemers_1.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers TblCustemers_2 ON dbo.TblAqar.ownerid = TblCustemers_2.CusID"
StrSQL = StrSQL & "   Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
    
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.BranchId = " & val(Me.DcbBranch.BoundText)
End If

If ChkAccredit.value = vbChecked Then
    StrWhere = StrWhere & " AND IsNull(TblContract.Accredit,0) = 1 "
End If

If ChkNotAccredit.value = vbChecked Then
    StrWhere = StrWhere & " AND IsNull(TblContract.Accredit,0) = 0 "
End If
If Not IsNull(XPDtbFrom.value) Then
    StrWhere = StrWhere & " AND  dbo.TblContract.ContDate >= " & SQLDate(XPDtbFrom.value, True) & ""
End If
If Not IsNull(XPDtpTo.value) Then
    StrWhere = StrWhere & " AND  dbo.TblContract.ContDate <= " & SQLDate(XPDtpTo.value, True) & ""
End If




If Me.dcaqartypeid.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid  = " & val(Me.dcaqartypeid.BoundText)
End If

If Me.DcbUnitNo.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqarDetai.ID = " & val(Me.DcbUnitNo.BoundText)
End If

If Me.dcbAqarType.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqarDetai.AqarID = " & val(Me.dcbAqarType.BoundText)
End If


If Me.dcsupplier.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.ownerid = " & val(dcsupplier.BoundText)
'gr = 2
End If

If Me.dbcClient.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqarDetai.customerid = " & val(dbcClient.BoundText)
'gr = 2
End If

'If Me.dcbSalesSpec.BoundText <> "" Then
'StrWhere = StrWhere & " AND cusid = " & val(dcbSalesSpec.BoundText)
'gr = 2
'End If


If Me.dcbCityId2.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CountryID = " & val(dcbCityId2.BoundText)
'gr = 2
End If

If Me.dcmCityID.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CityID = " & val(dcmCityID.BoundText)
'gr = 2
End If


If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqarDetai.unittype = " & val(DCAkarUnit.BoundText)
'gr = 2
End If


If Me.dbcAqarStatus.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblAqarDetai.Status = " & val(dbcAqarStatus.BoundText)
'gr = 2
End If


If Me.TxtStreet.Text <> "" Then
StrWhere = StrWhere & " AND  streetName like '%" & (TxtStreet.Text) & "%'"
'gr = 2
End If

If Me.txtRoomCount.Text <> "" Then
StrWhere = StrWhere & " AND  roomscount = " & val(txtRoomCount.Text)
'gr = 2
End If

If Me.XPDtbFrom.value = 1 Then
StrWhere = StrWhere & " AND  roomscount = " & val(txtRoomCount.Text)
'gr = 2
End If


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL, 6

    End If

End Sub

Public Sub GetData4()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
Dim s As String


s = " SELECT dbo.TblAqarDetai.unitno,"
s = s & "        dbo.TblAqar.Aqarid,TblAqar.aqarname,"
s = s & "       TblAkarUnit.Id        as   tblAkarTypeID,"
s = s & "       dbo.TblAqar.aqarname,TblAqarDetai.Id AS UnitID,"
s = s & "       dbo.TblAkarUnit.name   AS UnitName,"
s = s & "       dbo.TblAkarUnit.namee  AS UnitNamee,"


       
     '-- SELECT unitno From TblAqarDetai  where      Aqarid = " & ID & " and unittype=" & unittype
       
s = s & "       TotalW = ("
s = s & "           SELECT SUM(Net)"
        '   --,TblFiterWaiver.unittype,BulidID,ApartmentID,TblFiterWaiver.ID
s = s & "           From TblFiterWaiver"
        '   --GROUP BY TblFiterWaiver.unittype,BulidID,ApartmentID,Id
s = s & "           Where TblFiterWaiver.BulidID = TblAqar.Aqarid"
s = s & "                  AND TblFiterWaiver.unittype = TblAkarUnit.id"
s = s & "                  AND TblFiterWaiver.ApartmentID = TblAqarDetai.Id"
s = s & "       )"
s = s & " From dbo.TblAqarDetai"
s = s & "       INNER JOIN dbo.TblAqar"
s = s & "            ON  dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid"
s = s & "       LEFT OUTER JOIN dbo.tblAkarType"
s = s & "            ON  dbo.TblAqar.aqartypeid = dbo.tblAkarType.id"
s = s & "       LEFT OUTER JOIN dbo.TblContract"
s = s & "            ON  dbo.TblAqarDetai.ContID = dbo.TblContract.ContNo"
            
       
s = s & "       LEFT OUTER JOIN dbo.TblBranchesData"
s = s & "            ON  dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id"
s = s & "       LEFT OUTER JOIN dbo.TblRentStatus"
s = s & "            ON  dbo.TblAqarDetai.Status = dbo.TblRentStatus.id"
s = s & "       LEFT OUTER JOIN dbo.TblAkarUnit"
s = s & "            ON  dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id"



s = " SELECT"
'--ISNULL(TblContract.ContNo, 0)  AS Expr1,
s = s & "                   TblAqar.Aqarid, TblContract.UnitNo as UnitID,(SELECT top 1 TblAqarDetai.unitno FROM TblAqarDetai WHERE TblAqarDetai.Id = TblContract.UnitNo AND TblAqarDetai.Aqarid =TblAqar.Aqarid) as UnitNo,"
s = s & "                   TblAqar.aqarname,"
s = s & "                   TblAkarUnit.id                 AS tblAkarTypeID,"
s = s & "                   TblAqar.aqarname               AS aqarnamee,"
s = s & "                   TblAkarUnit.name               AS UnitName,"
s = s & "                   TblAkarUnit.namee              AS UnitNamee,"
s = s & "                   ("
s = s & "                       SELECT SUM(net) AS Expr1"
s = s & "                       From TblFiterWaiver"
s = s & "                       Where (BulidID = TblAqar.Aqarid)"
s = s & "                              AND (unittype        = TblContract.UnitType)"
s = s & "                              AND (ApartmentID     = TblContract.UnitNo)"
s = s & "                   )                              AS TotalW"
s = s & "            From TblAqar"
s = s & "                   INNER JOIN TblContract"
s = s & "                        ON  TblAqar.Aqarid = TblContract.Iqar"
s = s & "                   LEFT OUTER JOIN TblAkarUnit"
s = s & "                        ON  TblContract.UnitType = TblAkarUnit.id"
s = s & "                   LEFT OUTER JOIN tblAkarType"
s = s & "                        ON  TblAqar.aqartypeid = tblAkarType.id"
s = s & "                   LEFT OUTER JOIN TblBranchesData"
s = s & "                        ON  TblAqar.BranchId = TblBranchesData.branch_id"
s = s & "            Where (1 = 1)"




s = s & "       AND ISNULL(dbo.TblContract.ContNo, 0) <> 0"

If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.BranchId = " & val(Me.DcbBranch.BoundText)
End If

If ChkAccredit.value = vbChecked Then
    StrWhere = StrWhere & " AND IsNull(TblContract.Accredit,0) = 1 "
End If
If ChkNotAccredit.value = vbChecked Then
    StrWhere = StrWhere & " AND IsNull(TblContract.Accredit,0) = 0 "
End If

If Not IsNull(XPDtbFrom.value) Then
    StrWhere = StrWhere & " AND  dbo.TblContract.ContDate >= " & SQLDate(XPDtbFrom.value, True) & ""
End If
If Not IsNull(XPDtpTo.value) Then
    StrWhere = StrWhere & " AND  dbo.TblContract.ContDate <= " & SQLDate(XPDtpTo.value, True) & ""
End If



If Me.dcaqartypeid.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.aqartypeid  = " & val(Me.dcaqartypeid.BoundText)
End If

If Me.DcbUnitNo.BoundText <> "" Then
StrWhere = StrWhere & " AND TblContract.UnitNo= " & val(Me.DcbUnitNo.BoundText)
End If

If Me.dcbAqarType.BoundText <> "" Then
StrWhere = StrWhere & " AND TblContract.Iqar = " & val(Me.dcbAqarType.BoundText)
End If


If Me.dcsupplier.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.ownerid = " & val(dcsupplier.BoundText)
'gr = 2
End If

If Me.dbcClient.BoundText <> "" Then
StrWhere = StrWhere & " AND TblContract.CusID= " & val(dbcClient.BoundText)
'gr = 2
End If

'If Me.dcbSalesSpec.BoundText <> "" Then
'StrWhere = StrWhere & " AND cusid = " & val(dcbSalesSpec.BoundText)
'gr = 2
'End If


If Me.dcbCityId2.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CountryID = " & val(dcbCityId2.BoundText)
'gr = 2
End If

If Me.dcmCityID.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.CityID = " & val(dcmCityID.BoundText)
'gr = 2
End If


If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  TblAkarUnit.id= " & val(DCAkarUnit.BoundText)
'gr = 2
End If


'If Me.dbcAqarStatus.BoundText <> "" Then
'StrWhere = StrWhere & " AND  dbo.TblAqarDetai.Status = " & val(dbcAqarStatus.BoundText)
''gr = 2
'End If


If Me.TxtStreet.Text <> "" Then
StrWhere = StrWhere & " AND  streetName like '%" & (TxtStreet.Text) & "%'"
'gr = 2
End If

If Me.txtRoomCount.Text <> "" Then
StrWhere = StrWhere & " AND  roomscount = " & val(txtRoomCount.Text)
'gr = 2
End If

If Me.XPDtbFrom.value = 1 Then
StrWhere = StrWhere & " AND  roomscount = " & val(txtRoomCount.Text)
'gr = 2
End If


    '-----------------------------------
s = s & StrWhere


s = s & " Group By"
s = s & "        TblAqar.Aqarid,"
s = s & "        TblContract.UnitNo,"
s = s & "        TblAkarUnit.id,"
s = s & "        TblAqar.aqarname,"
s = s & "        TblAkarUnit.name,"
s = s & "        TblAkarUnit.namee,"
s = s & "        tblAkarType.id,"
s = s & "         TblContract.UnitType,"
s = s & "        TblAqar.aqartypeid"
s = s & " Order By"
s = s & "       dbo.TblAqar.Aqarid"


    BolBegine = False
    StrWhere = ""
    
    
 
  
  
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value


     
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    
    
    
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFilterWiaverTotal.rpt"
   

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    If ServersName.Text = "" Then
        mServer = ""
    Else
        mServer = ServersName.Text & ".dbo."
    End If
    
    
   
    
    
    'Dim s As String



s = " SELECT ISNULL(N.ID, Id1)             AS ID,"
s = s & "       ISNULL(NAME, Name1)          as    NAME,"
s = s & "        ISNULL(N.Aqarid, HH.Iqar1)       Aqarid,"
s = s & "        ISNULL(N.aqartypeid, HH.UnitType1) as tblAkarTypeID,"
s = s & "        ISNULL(N.unitno, HH.UnitNo1) as    UnitID,"
s = s & "        Total,"
s = s & "        HH.TotalItem,"
s = s & "        DiffT = IsNull(Total, 0) - IsNull(HH.TotalItem, 0)"
s = s & " FROM   ("
s = s & "            SELECT *"
s = s & "            FROM   ("
s = s & "                       SELECT SUM(TblFiterWaiverDe.Price * TblFiterWaiverDe.[Count]) AS Total,"
s = s & "                              TblAqrCompenet.ID,"
's = s & "                              --TblFiterWaiver.ID ddd,
s = s & "                              TblAqrCompenet.Name,"
s = s & "                              TblFiterWaiver.BulidID Aqarid,"
s = s & "                              TblFiterWaiver.unittype AS aqartypeid,"
s = s & "                              TblFiterWaiver.ApartmentID AS unitno"
s = s & "                       From TblFiterWaiverDe"
s = s & "                              LEFT OUTER JOIN dbo.TblFiterWaiver"
s = s & "                                   ON  TblFiterWaiver.ID = TblFiterWaiverDe.IDFItWaiv"
s = s & "                              LEFT OUTER JOIN dbo.TblAqrCompenetDet"
s = s & "                                   ON  dbo.TblFiterWaiverDe.IDItem = dbo.TblAqrCompenetDet.ID"
s = s & "                              LEFT OUTER JOIN TblAqrCompenet"
s = s & "                                   ON  TblAqrCompenetDet.IDAqComp = TblAqrCompenet.ID"
's = s & "                                       -- WHERE YEAR(TblFiterWaiver.FilterDate) = 2020
s = s & "                       Group By"
s = s & "                              TblAqrCompenet.ID,"
's = s & "                              --TblFiterWaiver.ID,
s = s & "                              TblAqrCompenet.Name,"
s = s & "                              TblFiterWaiver.BulidID,"
s = s & "                              TblFiterWaiver.unittype,"
s = s & "                              TblFiterWaiver.ApartmentID"
s = s & "                   ) T"
s = s & "            Where Total <> 0"
s = s & "        ) N"
s = s & "        FULL OUTER JOIN ("
s = s & "                 SELECT *"
s = s & "                 FROM   ("
s = s & "                            SELECT SUM(td.ShowQty * td.showPrice) AS TotalItem,"
s = s & "                                   TblAqrCompenet.ID ID1,"
s = s & "                                   TblAqrCompenet.Name Name1,"
s = s & "                                   t.Iqar Iqar1,"
s = s & "                                   t.UnitType UnitType1,"
s = s & "                                   t.unitno UnitNo1"
s = s & "                            FROM   " & mServer & "Transaction_Details AS td"
s = s & "                                   INNER JOIN " & mServer & "Transactions AS t"
s = s & "                                        ON  t.Transaction_ID = td.Transaction_ID"
s = s & "                                   INNER JOIN " & mServer & "TblItems AS ti"
s = s & "                                        ON  ti.ItemID = td.Item_ID"
s = s & "                                   LEFT OUTER JOIN " & mServer & "Groups AS g"
s = s & "                                        ON  g.GroupID = ti.GroupID"
s = s & "                                   LEFT OUTER JOIN " & mServer & "TblAqrCompenet"
s = s & "                                        ON  g.AqrCompenetId = TblAqrCompenet.ID"
s = s & "                            Where t.Transaction_Type = 21"
s = s & "                                   AND ISNULL(g.AqrCompenetId, 0) = TblAqrCompenet.ID"
s = s & "                                   AND ISNULL(t.Iqar, 0) <> 0"
s = s & "                            Group By"
s = s & "                                   TblAqrCompenet.ID,"
s = s & "                                   TblAqrCompenet.Name,"
s = s & "                                   t.Iqar,"
s = s & "                                   t.UnitType,"
s = s & "                                   t.unitno"
s = s & "                        )GG"
s = s & "             ) HH"
s = s & "             ON  N.ID = HH.ID1"
s = s & "             AND N.Aqarid = HH.Iqar1"
s = s & "             AND N.aqartypeid = HH.UnitType1"
s = s & "             AND N.unitno = HH.UnitNo1"
s = s & " Where n.Total <> 0"
s = s & "        OR  HH.TotalItem <> 0"
   Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    If s <> "" Then
        Dim RsData2  As New ADODB.Recordset
        
         
      '  RsData2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("aa").Database.SetDataSource RsData
  
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

    End If

End Sub


Function print_report2(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
      If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Aqar2.rpt"
       End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
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
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""

    End If
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
Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    If Rd(2) Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Aqar3ByAqar.rpt"
    ElseIf Rd(5) Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Aqar3ByBranch.rpt"
    Else
        If Ind = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportOtherExpenAqar.rpt"
           End If
        Else
          If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Aqar.rpt"
           End If
        End If
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

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

Private Sub Image3_Click()
AddTofaforites Me.Name, " ﬁ«—Ì— «·⁄ﬁ«—« ", " ﬁ«—Ì— «·⁄ﬁ«—« "
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Rd_Click(Index As Integer)
XPPnlTime.Visible = True
If Rd(1).value = True Then
XPPnlTime.Visible = True
End If
End Sub

Private Sub txtCodeCustomer_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeCustomer.Text, EmpID, , , 56
        dbcClient.BoundText = EmpID
    End If
End Sub

Private Sub txtCodeOwner_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeOwner.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
    End If
End Sub


Private Sub txtCodeOwner2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeOwner2.Text, EmpID, , , 57
        dcsupplier2.BoundText = EmpID
    End If
End Sub

Private Sub txtCodeSalesRep_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtCodeSalesRep.Text, EmpID
        dcbSalesSpec.BoundText = EmpID
    End If

End Sub


Private Sub txtRoomCount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, txtRoomCount)
End Sub




















Private Sub xpdtbfrom_Change()
If Not IsNull(XPDtbFrom.value) Then
XPDtbFromH.value = ToHijriDate(XPDtbFrom.value)
End If
End Sub

Private Sub XPDtbFromH_LostFocus()
   VBA.Calendar = vbCalGreg
            XPDtbFrom.value = ToGregorianDate(XPDtbFromH.value)
End Sub

Private Sub XPDtpTo_Change()
If Not IsNull(XPDtpTo.value) Then
XPDtpToH.value = ToHijriDate(XPDtpTo.value)
End If
End Sub

Private Sub XPDtpToH_LostFocus()
   VBA.Calendar = vbCalGreg
            XPDtpTo.value = ToGregorianDate(XPDtpToH.value)
End Sub
