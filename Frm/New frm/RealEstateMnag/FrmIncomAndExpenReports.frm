VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmIncomAndExpenReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "FrmIncomAndExpenReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   11640
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
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11655
      _cx             =   20558
      _cy             =   11033
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
      Caption         =   "ﬂ‘› «·Õ”«»|«·„’—Ê›«  Ê«·«Ì—«œ« "
      Align           =   0
      CurrTab         =   1
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
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5880
         Index           =   0
         Left            =   45
         TabIndex        =   35
         Top             =   45
         Width           =   11565
         Begin VB.CommandButton Command1 
            Caption         =   "„”Õ"
            Height          =   495
            Left            =   2640
            TabIndex        =   69
            Top             =   5280
            Width           =   1335
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ﬁ—Ì—"
            Height          =   615
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   720
            Width           =   4785
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Left            =   2400
               TabIndex        =   64
               Top             =   240
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«Ã„«·Ì «Ì—«œ Ê«·„’—Ê›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdAnalis 
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " Õ·Ì·Ì ··«Ì—«œ Ê«·„’—Ê›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   1455
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   2280
            Width           =   4695
            Begin MSComCtl2.DTPicker DtpDateFrom2 
               Height          =   330
               Left            =   1800
               TabIndex        =   48
               Top             =   150
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93257731
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo2 
               Height          =   330
               Left            =   1800
               TabIndex        =   49
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93257731
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromH2 
               Height          =   315
               Left            =   120
               TabIndex        =   50
               Top             =   150
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToH2 
               Height          =   315
               Left            =   120
               TabIndex        =   51
               Top             =   510
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   2
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   0
               Left            =   3270
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   600
               Width           =   480
            End
         End
         Begin VB.TextBox TxtSearch2 
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
            Left            =   3630
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ﬁ—Ì—"
            Height          =   615
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   720
            Width           =   6285
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   0
               Left            =   4200
               TabIndex        =   44
               Top             =   180
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «Ì—«œ«  Ê„’—Ê›« "
               ForeColor       =   12582912
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   3
               Left            =   2190
               TabIndex        =   45
               Top             =   180
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «·„’—Ê›« "
               ForeColor       =   12582912
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   4
               Left            =   60
               TabIndex        =   74
               Top             =   180
               Width           =   2115
               _Version        =   786432
               _ExtentX        =   3731
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— „—«Ã⁄… «·Õ—ﬂ… «·ÌÊ„Ì…"
               ForeColor       =   12582912
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   2280
            Width           =   5655
            Begin VB.ListBox SelectedTransTypeList 
               Height          =   1425
               ItemData        =   "FrmIncomAndExpenReports.frx":038A
               Left            =   120
               List            =   "FrmIncomAndExpenReports.frx":0391
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   360
               Width           =   2325
            End
            Begin VB.ListBox TransTypeList 
               Height          =   1425
               ItemData        =   "FrmIncomAndExpenReports.frx":03A4
               Left            =   3090
               List            =   "FrmIncomAndExpenReports.frx":03AB
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   360
               Width           =   2370
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   510
               Width           =   480
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   825
               Width           =   375
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   1170
               Width           =   570
            End
         End
         Begin MSDataListLib.DataCombo DcbBranch2 
            Height          =   315
            Left            =   5760
            TabIndex        =   54
            Top             =   1440
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbIqara2 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
            Top             =   1440
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitNo2 
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   1800
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitType2 
            Height          =   315
            Left            =   5760
            TabIndex        =   57
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   1800
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   1320
            TabIndex        =   70
            Top             =   5280
            Width           =   1245
            _ExtentX        =   2196
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
            Top             =   5280
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
         Begin VB.Image Image1 
            Height          =   390
            Left            =   3720
            Picture         =   "FrmIncomAndExpenReports.frx":03B8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "‘«‘…  ﬁ—Ì— »«·«Ì—«œ Ê«·„’—Ê›"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   -45
            TabIndex        =   72
            Top             =   0
            Width           =   11640
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   195
            Index           =   10
            Left            =   10995
            TabIndex        =   62
            Top             =   1320
            Width           =   345
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   615
            Left            =   120
            Top             =   4560
            Width           =   11295
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »≈ŸÂ«—  «·«Ì—«œ Ê«·„’—Ê› ÿ»ﬁ« ·· √—ÌŒ"
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
            Height          =   540
            Index           =   5
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   4560
            Width           =   11175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÊÕœ…"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·ÊÕœ…"
            Height          =   195
            Index           =   8
            Left            =   10350
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1800
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄ﬁ«—"
            Height          =   195
            Index           =   7
            Left            =   4545
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1440
            Width           =   990
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5880
         Index           =   1
         Left            =   -12210
         TabIndex        =   5
         Top             =   45
         Width           =   11565
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Left            =   2760
            TabIndex        =   32
            Top             =   5280
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ﬁ—Ì—"
            Height          =   615
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   600
            Width           =   5295
            Begin XtremeSuiteControls.RadioButton RdRep 
               Height          =   255
               Index           =   0
               Left            =   3480
               TabIndex        =   66
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„’—Ê›«  ›ﬁÿ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdRep 
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   67
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«Ì—«œ«  ›ﬁÿ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdRep 
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   68
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·ﬂ·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   735
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   2640
            Width           =   10575
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   6600
               TabIndex        =   12
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93257731
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   1800
               TabIndex        =   13
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93257731
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
               Height          =   315
               Left            =   4920
               TabIndex        =   14
               Top             =   270
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToH 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   270
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   4
               Left            =   8850
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   3
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.TextBox TxtSearch 
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
            Left            =   3630
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ﬁ—Ì—"
            Height          =   615
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   5775
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   1
               Left            =   2520
               TabIndex        =   8
               Top             =   120
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "ﬂ‘› Õ”«» ⁄ﬁ«—"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "ﬂ‘› Õ”«» "
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3630
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1065
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   6000
            TabIndex        =   18
            Top             =   1560
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbIqara 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
            Top             =   1920
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitNo 
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   2280
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUnitType 
            Height          =   315
            Left            =   6000
            TabIndex        =   21
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   2280
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbAccount 
            Height          =   315
            Left            =   6000
            TabIndex        =   22
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
            Top             =   1920
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountsDC 
            Bindings        =   "FrmIncomAndExpenReports.frx":4020
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   1560
            Width           =   3435
            _ExtentX        =   6059
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   1440
            TabIndex        =   33
            Top             =   5280
            Width           =   1245
            _ExtentX        =   2196
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
            Left            =   240
            TabIndex        =   34
            Top             =   5280
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
            Caption         =   "‘«‘… ﬂ‘› «·Õ”«»"
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
            Left            =   -15
            TabIndex        =   73
            Top             =   0
            Width           =   11610
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   195
            Index           =   0
            Left            =   10440
            TabIndex        =   30
            Top             =   1560
            Width           =   990
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   1695
            Left            =   120
            Top             =   3480
            Width           =   11295
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »≈ŸÂ«—  ﬂ‘› «·Õ”«»  ÿ»ﬁ« ·· √—ÌŒ"
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
            Height          =   1620
            Index           =   25
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   3480
            Width           =   11295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÊÕœ…"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   14
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   2280
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·ÊÕœ…"
            Height          =   195
            Index           =   15
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2280
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄ﬁ«—"
            Height          =   195
            Index           =   4
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1920
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„’—Ê›"
            Height          =   195
            Index           =   1
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1920
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«»"
            Height          =   195
            Index           =   2
            Left            =   4800
            TabIndex        =   24
            Top             =   1560
            Width           =   870
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
      _cx             =   1931
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
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   600
      Picture         =   "FrmIncomAndExpenReports.frx":4035
      Stretch         =   -1  'True
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   2
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmIncomAndExpenReports"
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
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub AccountsDC_Change()
AccountsDC_Click (0)
End Sub

Private Sub AccountsDC_Click(Area As Integer)
    Dim ClsAcc As New ClsAccounts
    Set ClsAcc = New ClsAccounts
    Text1.Text = ClsAcc.Get_Account_Serial(AccountsDC.BoundText)
End Sub

Private Sub AccountsDC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 889
    End If
End Sub

Private Sub Command1_Click()
clear_all Me
Me.RdAnalis.value = False
Me.RdTotal.value = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
Me.SelectedTransTypeList.Clear
End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
If DcbAccount.BoundText <> "" Then
AccountsDC.BoundText = DcbAccount.BoundText
End If
End Sub

Private Sub DcbIqara_Change()
DcbUnitType_Change
DcbIqara_Click (0)
End Sub
Public Function ReloadCombos()
Dim Dcombos As ClsDataCombos
 
 Set Dcombos = New ClsDataCombos
   Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetExpensesAccount Me.DcbAccount
    Dcombos.GetBranches DcbBranch
   End Function

Private Sub DcbIqara2_Change()
DcbIqara2_Click (0)
DcbUnitType2_Click (0)
End Sub

Private Sub DcbIqara2_Click(Area As Integer)
     If val(DcbIqara2.BoundText) = 0 Then: Exit Sub
     Dim EmpCode  As String
     Dim ownerid As Double
    GetIqarCode , , DcbIqara2.BoundText, EmpCode, ownerid
    Me.TxtSearch2.Text = EmpCode
End Sub

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos
If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)
idd1 = val(DcbUnitType.BoundText)
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
End If
End Sub
Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub
Private Sub DcbUnitType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub
Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub
     Dim EmpCode  As String
     Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    Me.TxtSearch.Text = EmpCode
End Sub

Private Sub DcbUnitType2_Change()
DcbUnitType2_Click (0)
End Sub

Private Sub DcbUnitType2_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos
If val(DcbIqara2.BoundText) > 0 Then
idd = val(DcbIqara2.BoundText)
idd1 = val(DcbUnitType2.BoundText)
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo2, "R"
End If
End Sub

Private Sub DtpDateFrom2_Change()
If Not IsNull(DtpDateFrom2.value) Then
   DtpDateFromh2.value = ToHijriDate(DtpDateFrom2.value)
   End If
End Sub

Private Sub DtpDateFromh2_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom2.value = ToGregorianDate(DtpDateFromh2.value)
End Sub

Private Sub DtpDateTo2_Change()
If Not IsNull(DtpDateTo2.value) Then
   DtpDateToh2.value = ToHijriDate(DtpDateTo2.value)
   End If
End Sub

Private Sub DtpDateToh2_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo2.value = ToGregorianDate(DtpDateToh2.value)
End Sub

Private Sub Image1_Click()
AddTofaforites Me.Name, " ﬁ«—Ì— «Ì—«œ Ê„’—Ê›", " ﬁ«—Ì— «Ì—«œ Ê„’—Ê›"
End Sub

Sub account()
 
            ShowReport AccountsDC.BoundText, AccountsDC.Text, DtpDateFrom.value, DtpDateTo.value, , val(DcbBranch.BoundText)
End Sub


Private Sub Label12_Click()
    If Me.SelectedTransTypeList.ListIndex > -1 Then
        Me.SelectedTransTypeList.RemoveItem (SelectedTransTypeList.ListIndex)
    End If
End Sub

Private Sub Label13_Click()
 Me.SelectedTransTypeList.Clear
End Sub
Function GetTransIds() As String
    Dim tempString As String
    Dim i As Integer
    tempString = "0"
    For i = 0 To Me.SelectedTransTypeList.ListCount - 1
        tempString = tempString & "," & Me.SelectedTransTypeList.ItemData(i)
    Next i
    GetTransIds = tempString
End Function
Private Sub Label14_Click()
    Dim i As Integer
    Me.SelectedTransTypeList.Clear
    For i = 0 To Me.TransTypeList.ListCount - 1
        Me.SelectedTransTypeList.AddItem TransTypeList.List(i)
        SelectedTransTypeList.ItemData(i) = TransTypeList.ItemData(i)
    Next i
End Sub

Private Sub Label15_Click()
    If Me.TransTypeList.ListIndex > -1 Then
        Me.SelectedTransTypeList.AddItem TransTypeList.List(TransTypeList.ListIndex)
        SelectedTransTypeList.ItemData(SelectedTransTypeList.NewIndex) = TransTypeList.ItemData(TransTypeList.ListIndex)
    End If
End Sub


Private Sub Rd_Click(Index As Integer)
If Rd(1).value = True Then
TxtSearch.Enabled = True
DcbIqara.Enabled = True
DcbUnitNo.Enabled = True
DcbUnitType.Enabled = True
ElseIf Rd(2).value = True Then
TxtSearch.Enabled = False
DcbIqara.Enabled = False
DcbUnitNo.Enabled = False
DcbUnitType.Enabled = False
End If
End Sub

Private Sub RdAnalis_Click()
TxtSearch.Enabled = True
DcbIqara.Enabled = True
DcbUnitType.Enabled = True
DcbUnitNo.Enabled = True
DcbAccount.Enabled = True
End Sub

Private Sub RdTotal_Click()
TxtSearch.Enabled = False
DcbIqara.Enabled = False
DcbUnitType.Enabled = False
DcbUnitNo.Enabled = False
DcbAccount.Enabled = False
End Sub
Public Function GetAccountCode(StrAccSerial As String) As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    If Trim(StrAccSerial) <> "" Then
        If Trim(StrAccSerial) = "" Then Exit Function
        StrSQL = "Select Account_Code From ACCOUNTS Where Account_Serial ='" & Trim(StrAccSerial) & "'"
        
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetAccountCode = rs("Account_Code").value
        Else
        End If

        rs.Close
        Set rs = Nothing
    End If
End Function
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If Text1.Text = "" Then
            Me.AccountsDC.BoundText = ""
        Else
            Me.AccountsDC.BoundText = GetAccountCode(Trim$(Me.Text1.Text))
        End If
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 889
    End If
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub
Private Sub btnClear_Click()
clear_all Me
Me.RdAnalis.value = False
Me.RdTotal.value = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
Me.SelectedTransTypeList.Clear
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
        If Rd(1).value = False And Rd(2).value = False Then
        MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄ «· ﬁ—Ì—"
        Exit Sub
        End If
'If Rd(0).value = True Or Rd(3).value = True Then
'If Me.RdAnalis.value = False And Me.RdTotal.value = False And Rd(3).value = False Then
'MsgBox "Ì—ÃÏ «Œ «Ì— ‰Ê⁄ «· ﬁ—Ì— «Ã„«·Ì «Ê  Õ·Ì·Ì"
'Exit Sub
'End If
'If Me.RdAnalis.value = True Or Me.RdTotal.value = True Or Rd(3).value = True Then
'GetData
'End If
'End If

If Rd(1).value = True Then
If val(DcbIqara.BoundText) = 0 Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄ﬁ«—"
DcbIqara.SetFocus
Exit Sub
End If
ShowGLto_Aqar val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)
End If

If Rd(2).value = True Then
If (AccountsDC.BoundText) = "" Or AccountsDC.Text = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
AccountsDC.SetFocus
Exit Sub
End If
If IsNull(DtpDateFrom.value) Or IsNull(DtpDateTo.value) Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·› —…"
Exit Sub
End If
account
End If
        Case 1
            clear_all Me

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2, 3
            Unload Me
        Case 4
           If Rd(4) Then
        GetData3
        Exit Sub
    End If
        If Rd(0).value = False And Rd(3).value = False Then
        MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄ «· ﬁ—Ì—"
        Exit Sub
        End If
   
If Rd(0).value = True Or Rd(3).value = True Then
If Me.RdAnalis.value = False And Me.RdTotal.value = False And Rd(3).value = False Then
MsgBox "Ì—ÃÏ «Œ «Ì— ‰Ê⁄ «· ﬁ—Ì— «Ã„«·Ì «Ê  Õ·Ì·Ì"
Exit Sub
End If
If Me.RdAnalis.value = True Or Me.RdTotal.value = True Or Rd(3).value = True Then
GetData
ElseIf Rd(4) Then
    GetData3
End If
End If
        
'print_report
    End Select

End Sub
Private Sub DtpDateFrom_Change()
If Not IsNull(DtpDateFrom.value) Then
   DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
   End If
End Sub
Private Sub DtpDateFromH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub
Private Sub DtpDateTo_Change()
If Not IsNull(DtpDateTo.value) Then
   DtpDateToH.value = ToHijriDate(DtpDateTo.value)
   End If
End Sub


Private Sub DtpDateToH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub

Function ShowGLto_Aqar(Aqarid As Double, Optional unittype As Double, Optional unitno As Double)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim NotesTypes As String
    Dim StrFileName As String
    NotesTypes = ""
    Dim Msg As String
    MySQL = "  SELECT     dbo.Notes.NoteID, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.Notes.CashingType, dbo.Notes.UnitNo, TblAqar_1.aqarname,"
    MySQL = MySQL & "                   dbo.Notes.akarid, dbo.Notes.branch_no, TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, TblBranchesData_1.branch_id,"
    MySQL = MySQL & "                   TblAqarDetai_2.unittype, TblAkarUnit_2.name, TblAkarUnit_2.namee, TblAqarDetai_2.unitno AS Nameunitno, TblAqarDetai_2.Aqarid, TblAqarDetai_2.Id,"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
    MySQL = MySQL & "                   TblAqar_2.aqarname AS aqarnameDet, dbo.DOUBLE_ENTREY_VOUCHERS.unittype AS unittypeDet, TblAkarUnit_1.name AS nameDob,"
    MySQL = MySQL & "                   TblAkarUnit_1.namee AS nameDobE, dbo.DOUBLE_ENTREY_VOUCHERS.unitno AS unitnoDet, TblAqarDetai_1.unitno AS unitnoDetE, dbo.Notes.NoteDateH,"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.notes_all.branch_no AS branch_no1, TblBranchesData_1.branch_name AS branch_name1,"
    MySQL = MySQL & "                   TblBranchesData_1.branch_namee AS branch_namee1, TblBranchesData_1.branch_id AS branch_id1, dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid AS Expr1,"
    MySQL = MySQL & "                   dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DEV_DES,"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.Notes.NoteDate, dbo.Notes.NoteSerial1, TblAqar_1.aqarNo,"
    MySQL = MySQL & "                   TblAqar_2.aqarNo AS aqarNo2 ,dbo.ACCOUNTS.Account_Serial"
    MySQL = MySQL & "      FROM         dbo.TblNotesTypes INNER JOIN"
    MySQL = MySQL & "                   dbo.Notes ON dbo.TblNotesTypes.NotesType = dbo.Notes.NoteType RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.notes_all ON TblBranchesData_1.branch_id = dbo.notes_all.branch_no RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.notes_all.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_all ON"
    MySQL = MySQL & "                   dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqarDetai TblAqarDetai_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unitno = TblAqarDetai_1.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAkarUnit TblAkarUnit_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqar TblAqar_2 ON dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid = TblAqar_2.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqarDetai TblAqarDetai_2 LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAqar TblAqar_1 ON TblAqarDetai_2.Aqarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAkarUnit TblAkarUnit_2 ON TblAqarDetai_2.unittype = TblAkarUnit_2.id ON dbo.Notes.UnitNo = TblAqarDetai_2.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData TblBranchesData_2 ON dbo.Notes.branch_no = TblBranchesData_2.branch_id"
    MySQL = MySQL & "      WHERE   (  (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0))  "
    'and ((dbo.Notes.NoteType=3  and dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid =" & Aqarid & ") or (dbo.Notes.NoteType=5  and dbo.Notes.akarid =" & Aqarid & ")))"
   'MySQL = MySQL & "      or     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1 and (dbo.Notes.NoteType=350)) and dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid =" & Aqarid & ")"
    If Aqarid <> 0 Then
      MySQL = MySQL & " AND ((ISNULL(dbo.Notes.akarid,ISNULL(dbo.Notes.akarid,0)) = " & Aqarid & ") or (ISNULL(dbo.Notes.akarid,ISNULL(dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid,0)) = " & Aqarid & ")) "
       'OR ( = " & Aqarid & "))"
    End If

    If val(Me.DcbBranch.BoundText) <> 0 And DcbBranch.Text <> "" Then
    MySQL = MySQL & " and  dbo.DOUBLE_ENTREY_VOUCHERS.branch_id=" & val(DcbBranch.BoundText) & " "
    End If
    NotesTypes = "0"
   If RdRep(0).value = True Then
    NotesTypes = "3,5,350"
    ElseIf RdRep(1).value = True Then
    NotesTypes = "4"
    ElseIf RdRep(2).value = True Then
    NotesTypes = "0"
   End If
    
    If NotesTypes <> "0" Then
    MySQL = MySQL & " and  dbo.Notes.NoteType in (" & NotesTypes & ") "
    End If
   If AccountsDC.Text <> "" And AccountsDC.BoundText <> "" Then
   MySQL = MySQL & " and  dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code='" & AccountsDC.BoundText & "'"
   End If
 
    If unittype <> 0 Then
        MySQL = MySQL & " and ( ( ISNULL(dbo.Notes.unittype,0)=" & unittype & " ) or (ISNULL(dbo.DOUBLE_ENTREY_VOUCHERS.unittype,0)=" & unittype & ")) "
    End If
       If unitno <> 0 Then
       MySQL = MySQL & "  and (( ISNULL(dbo.Notes.UnitNo,0)=" & unitno & ") or ( ISNULL(dbo.DOUBLE_ENTREY_VOUCHERS.unitno,0)=" & unitno & ")) "
    End If
    


    If Not IsNull(Me.DtpDateFrom.value) Then
        MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        MySQL = MySQL + " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
    If RdRep(0).value = True Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "GL _with_Aqar.rpt"
    ElseIf RdRep(1).value = True Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "GL _with_Aqar1.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "GL _with_AqarAll.rpt"
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
            Msg = "No data to view"
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
       ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic


    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        StrReportTitle = ""

    End If
    If RdRep(2).value = True Then
        StrReportTitle = "ﬂ‘› Õ”«» «·⁄ﬁ«— " '& StrAccountName

        If unittype <> 0 Then
                StrReportTitle = " ﬂ‘› Õ”«» «·⁄ﬁ«— " + DcbIqara.Text + "«·‰Ê⁄ " + Me.DcbUnitType.Text
        End If
        If unitno <> 0 Then
                StrReportTitle = " ﬂ‘› Õ”«» «·⁄ﬁ«— " + DcbIqara.Text + "«·‰Ê⁄ " + Me.DcbUnitType.Text + "—ﬁ„ «·ÊÕœ… " + Me.DcbUnitNo.Text
        End If

        If Me.DtpDateFrom.value <> Empty Or Me.DtpDateFrom.value <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DtpDateFrom.value, "yyyy/M/d") & ""
        End If

        If Me.DtpDateTo.value <> Empty Or Me.DtpDateTo.value <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DtpDateTo.value, "yyyy/M/d") & " "
        End If
    ElseIf RdRep(1).value = True Then
    StrReportTitle = "«Ì—«œ« "
    ElseIf RdRep(0).value = True Then
    StrReportTitle = "„’—Ê›« "
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue TxtSearch.Text
    xReport.ParameterFields(5).AddCurrentValue DcbIqara.Text
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


Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetExpensesAccount Me.DcbAccount
    Dcombos.GetBranches DcbBranch
    Dcombos.GetBranches DcbBranch2
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqar DcbIqara2
    Dcombos.getAkarUnit Me.DcbUnitType2
    'Dcombos.GetRentStatus dbcAqarStatus
Me.RdAnalis.value = False
Me.RdTotal.value = False
DtpDateFrom.value = Date
DtpDateTo.value = Date
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = Date
DtpDateTo2.value = Date
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
TransTypeList.Clear
SelectedTransTypeList.Clear
Dcombos.GetAccountingCodes Me.AccountsDC
FillLists
    Resize_Form Me
End Sub

Sub FillLists()
    Dim listRS As ADODB.Recordset
    Set listRS = New ADODB.Recordset
    Dim i As Integer
    Dim listSQL As String
    '--------------------------------------------------------------------------------------------------------------------------------------------

    '--------------------------------------------------------------------------------------------------------------------------------------------
    listSQL = "select * from TblNotesTypes where NotesType in (3,4,5,350)"
    listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    TransTypeList.AddItem IIf(IsNull(listRS("NotesTypeName").value), "", listRS("NotesTypeName").value)
                Else
                    TransTypeList.AddItem IIf(IsNull(listRS("NotesTypeNamee").value), "", listRS("NotesTypeNamee").value)
                End If
                TransTypeList.ItemData(TransTypeList.NewIndex) = IIf(IsNull(listRS("NotesType").value), 0, listRS("NotesType").value)
                listRS.MoveNext
            Next i
        End If
    listRS.Close
    End Sub


Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

If Me.RdTotal.value = True Then

StrSQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteDate, SUM(dbo.Notes.Note_Value) AS total, dbo.Notes.branch_no, dbo.TblBranchesData.branch_id, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE , dbo.Notes.NoteDateH "
StrSQL = StrSQL & " FROM         dbo.Notes LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = 3 OR"
StrSQL = StrSQL & "                      dbo.Notes.NoteType = 4 OR"
StrSQL = StrSQL & "                      dbo.Notes.NoteType = 5 or dbo.Notes.NoteType = 350)"
    If GetTransIds <> "0" Then
    StrSQL = StrSQL & " and  dbo.Notes.NoteType in (" & GetTransIds & ") "
    End If

If Me.DcbUnitType2.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.unittype = " & val(Me.DcbUnitType2.BoundText)

End If
If Me.DcbUnitNo2.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.UnitNo = " & val(Me.DcbUnitNo2.BoundText)

End If

End If
If Me.RdAnalis.value = True Or Rd(3).value = True Then
StrSQL = " SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.CashingType, "
StrSQL = StrSQL & "                      dbo.Notes.UnitNo, TblAqar_1.aqarname, dbo.Notes.akarid, dbo.Notes.branch_no, TblAqarDetai_2.unittype, TblAkarUnit_2.name, TblAkarUnit_2.namee,"
StrSQL = StrSQL & "                      TblAqarDetai_2.unitno AS Nameunitno, TblAqarDetai_2.Aqarid, TblAqarDetai_2.Id, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name,"
StrSQL = StrSQL & "                      dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid AS AqaridDet, TblAqar_2.aqarname AS aqarnameDet, dbo.DOUBLE_ENTREY_VOUCHERS.unittype AS unittypeDet,"
StrSQL = StrSQL & "                      TblAkarUnit_1.name AS nameDob, TblAkarUnit_1.namee AS nameDobE, dbo.DOUBLE_ENTREY_VOUCHERS.unitno AS unitnoDet, TblAqarDetai_1.unitno AS unitnoDetE,"
StrSQL = StrSQL & "                       dbo.Notes.NoteDateH, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.notes_all.branch_no AS branch_no1, dbo.Notes.BankID, BanksData_1.BankName,"
StrSQL = StrSQL & "                      dbo.Notes.BoxID, TblBoxesData_1.BoxName, BanksData_1.BankName AS BankName2, TblBoxesData_1.BoxName AS BoxName2,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, TblBranchesData_1.branch_name AS branch_name2,"
StrSQL = StrSQL & "                      TblBranchesData_1.branch_namee AS branch_name2E , dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " FROM         dbo.BanksData BanksData_2 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Notes ON TblBranchesData_1.branch_id = dbo.Notes.branch_no LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBoxesData TblBoxesData_2 ON dbo.Notes.BoxID = TblBoxesData_2.BoxID ON BanksData_2.BankID = dbo.Notes.BankID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.BanksData BanksData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.notes_all ON dbo.TblBranchesData.branch_id = dbo.notes_all.branch_no RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.notes_all.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_all LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBoxesData TblBoxesData_1 ON dbo.notes_all.BoxID = TblBoxesData_1.BoxID ON BanksData_1.BankID = dbo.notes_all.BankID ON"
StrSQL = StrSQL & "                      dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai TblAqarDetai_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unitno = TblAqarDetai_1.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit TblAkarUnit_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar TblAqar_2 ON dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid = TblAqar_2.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai TblAqarDetai_2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar TblAqar_1 ON TblAqarDetai_2.Aqarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit TblAkarUnit_2 ON TblAqarDetai_2.unittype = TblAkarUnit_2.id ON dbo.Notes.UnitNo = TblAqarDetai_2.Id"
StrSQL = StrSQL & " where 1=1"
If Rd(3).value = True Then
StrSQL = StrSQL & "  and      (dbo.Notes.NoteType = 3 OR"
StrSQL = StrSQL & "                       dbo.Notes.NoteType = 5 or  dbo.Notes.NoteType = 350) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
Else
StrSQL = StrSQL & "  and      (dbo.Notes.NoteType = 3 OR"
StrSQL = StrSQL & "                       dbo.Notes.NoteType = 4 OR"
StrSQL = StrSQL & "                       dbo.Notes.NoteType = 5 or  dbo.Notes.NoteType = 350) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
End If
If Me.DcbIqara2.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.akarid = " & val(Me.DcbIqara2.BoundText)
End If
    If GetTransIds <> "0" And Rd(3).value = False Then
    StrSQL = StrSQL & " and  dbo.Notes.NoteType in (" & GetTransIds & ") "
    End If


If Me.DcbUnitType2.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.unittype = " & val(Me.DcbUnitType2.BoundText)

End If
If Me.DcbUnitNo2.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.UnitNo = " & val(Me.DcbUnitNo2.BoundText)

End If

End If

'''''''''
 If Not IsNull(Me.DtpDateFrom2.value) Then
                   StrSQL = StrSQL & " AND dbo.Notes.NoteDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo2.value) Then
                   StrSQL = StrSQL & " AND dbo.Notes.NoteDate<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
      End If
      
If Me.DcbBranch2.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch2.BoundText)
End If

If Me.RdTotal.value = True And Rd(3).value = False Then
StrSQL = StrSQL & " GROUP BY dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.branch_no, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_nameE , dbo.Notes.NoteHijriDate , dbo.Notes.NoteDateH"

End If


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
    End If

End Sub
Public Sub GetData2()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

If Me.RdTotal.value = True Then

StrSQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteDate, SUM(dbo.Notes.Note_Value) AS total, dbo.Notes.branch_no, dbo.TblBranchesData.branch_id, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE , dbo.Notes.NoteDateH "
StrSQL = StrSQL & " FROM         dbo.Notes LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = 3 OR"
StrSQL = StrSQL & "                      dbo.Notes.NoteType = 4 OR"
StrSQL = StrSQL & "                      dbo.Notes.NoteType = 5 or dbo.Notes.NoteType = 350)"
    If GetTransIds <> "0" Then
    StrSQL = StrSQL & " and  dbo.Notes.NoteType in (" & GetTransIds & ") "
    End If

End If
If Me.RdAnalis.value = True Or Rd(3).value = True Then
StrSQL = " SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.CashingType, "
StrSQL = StrSQL & "                      dbo.Notes.UnitNo, TblAqar_1.aqarname, dbo.Notes.akarid, dbo.Notes.branch_no, TblAqarDetai_2.unittype, TblAkarUnit_2.name, TblAkarUnit_2.namee,"
StrSQL = StrSQL & "                      TblAqarDetai_2.unitno AS Nameunitno, TblAqarDetai_2.Aqarid, TblAqarDetai_2.Id, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name,"
StrSQL = StrSQL & "                      dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid AS AqaridDet, TblAqar_2.aqarname AS aqarnameDet, dbo.DOUBLE_ENTREY_VOUCHERS.unittype AS unittypeDet,"
StrSQL = StrSQL & "                      TblAkarUnit_1.name AS nameDob, TblAkarUnit_1.namee AS nameDobE, dbo.DOUBLE_ENTREY_VOUCHERS.unitno AS unitnoDet, TblAqarDetai_1.unitno AS unitnoDetE,"
StrSQL = StrSQL & "                       dbo.Notes.NoteDateH, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.notes_all.branch_no AS branch_no1, dbo.Notes.BankID, BanksData_1.BankName,"
StrSQL = StrSQL & "                      dbo.Notes.BoxID, TblBoxesData_1.BoxName, BanksData_1.BankName AS BankName2, TblBoxesData_1.BoxName AS BoxName2,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, TblBranchesData_1.branch_name AS branch_name2,"
StrSQL = StrSQL & "                      TblBranchesData_1.branch_namee AS branch_name2E , dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " FROM         dbo.BanksData BanksData_2 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Notes ON TblBranchesData_1.branch_id = dbo.Notes.branch_no LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBoxesData TblBoxesData_2 ON dbo.Notes.BoxID = TblBoxesData_2.BoxID ON BanksData_2.BankID = dbo.Notes.BankID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.BanksData BanksData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.notes_all ON dbo.TblBranchesData.branch_id = dbo.notes_all.branch_no RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.notes_all.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_all LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBoxesData TblBoxesData_1 ON dbo.notes_all.BoxID = TblBoxesData_1.BoxID ON BanksData_1.BankID = dbo.notes_all.BankID ON"
StrSQL = StrSQL & "                      dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai TblAqarDetai_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unitno = TblAqarDetai_1.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit TblAkarUnit_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar TblAqar_2 ON dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid = TblAqar_2.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai TblAqarDetai_2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar TblAqar_1 ON TblAqarDetai_2.Aqarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit TblAkarUnit_2 ON TblAqarDetai_2.unittype = TblAkarUnit_2.id ON dbo.Notes.UnitNo = TblAqarDetai_2.Id"
StrSQL = StrSQL & " where 1=1"
If Rd(3).value = True Then
StrSQL = StrSQL & "  and      (dbo.Notes.NoteType = 3 OR"
StrSQL = StrSQL & "                       dbo.Notes.NoteType = 5 or  dbo.Notes.NoteType = 350) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
Else
StrSQL = StrSQL & "  and      (dbo.Notes.NoteType = 3 OR"
StrSQL = StrSQL & "                       dbo.Notes.NoteType = 4 OR"
StrSQL = StrSQL & "                       dbo.Notes.NoteType = 5 or  dbo.Notes.NoteType = 350) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
End If
If Me.DcbIqara.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.akarid = " & val(Me.DcbIqara.BoundText)
End If
    If GetTransIds <> "0" And Rd(3).value = False Then
    StrSQL = StrSQL & " and  dbo.Notes.NoteType in (" & GetTransIds & ") "
    End If
   If AccountsDC.Text <> "" And AccountsDC.BoundText <> "" Then
   StrSQL = StrSQL & " and  dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code='" & AccountsDC.BoundText & "'"
   End If
If Me.DcbAccount.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = '" & Me.DcbAccount.BoundText & " '"

End If

If Me.DcbUnitType.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.unittype = " & val(Me.DcbUnitType.BoundText)

End If
If Me.DcbUnitNo.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.Notes.UnitNo = " & val(Me.DcbUnitNo.BoundText)

End If

End If

'''''''''
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Notes.NoteDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Notes.NoteDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
      
If Me.DcbBranch.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)
End If

If Me.RdTotal.value = True And Rd(3).value = False Then
StrSQL = StrSQL & " GROUP BY dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.branch_no, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_nameE , dbo.Notes.NoteHijriDate , dbo.Notes.NoteDateH"

End If


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
    End If

End Sub

Public Sub GetData3()
    Dim MySQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer




MySQL = " SELECT     dbo.Notes.NoteID,tu.UserName, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, IsNull(dbo.Notes.Note_Value2,0) + IsNull(Notes.Vat,0) Note_Value, dbo.Notes.NoteDateH,"
MySQL = MySQL & "                       dbo.Notes.ContractNo, dbo.Notes.ContNo, dbo.Notes.commission, dbo.Notes.rent, dbo.Notes.Water, dbo.Notes.FilterID, dbo.Notes.FIlterTotal, dbo.Notes.Instrunce,"
MySQL = MySQL & "                       dbo.Notes.comX, dbo.Notes.ComY, dbo.Notes.CommissionOut, dbo.Notes.NoteOrBonID, dbo.Notes.comXold, dbo.Notes.ComYold, dbo.Notes.NoteOrBonValue,"
MySQL = MySQL & "                       dbo.Notes.NoteOrBonSereal, dbo.Notes.Telephone, dbo.Notes.CashingType, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "                       dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.Notes.renterName, dbo.Notes.NoteCashingType, dbo.Notes.BankName, dbo.Notes.DueDate,"
MySQL = MySQL & "                       dbo.Notes.ChqueNum, dbo.Notes.Remark, dbo.Notes.Remark2, dbo.Notes.ToPriodDateH, dbo.Notes.FrmPriodDateH, dbo.Notes.ToPriodDate, dbo.Notes.FrmPriodDate,"
MySQL = MySQL & "                       dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.unitno,"
MySQL = MySQL & "                       dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.Aqarid, TblAqar_1.aqarname, TblAkarUnit_2.name, TblAkarUnit_2.namee, dbo.Notes.akarid,"
                      MySQL = MySQL & " TblAqar_1.aqarname AS aqarname2, dbo.Notes.unittype AS unittype2, TblAkarUnit_1.name AS name2, TblAkarUnit_1.namee AS namee2, dbo.Notes.Electricity,"
MySQL = MySQL & "                       dbo.Notes.BankID, dbo.BanksData.BankNamee, dbo.BanksData.BankName AS BankName2, dbo.Notes.Servce"
', dbo.TblNotesSales.rate, dbo.TblNotesSales.valu,"
'MySQL = MySQL & "                       dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
' MySQL = MySQL & "                      dbo.Notes.RemaiValue, dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.RentValuePayed,"

'MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.CommissionsPayed, dbo.ContracttBillInstallmentsDone.InsurancePayed, dbo.ContracttBillInstallmentsDone.ElectricPayed,"
'MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.TelandNetPayed, dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH,"
'MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.total, dbo.ContracttBillInstallmentsDone.[Value], dbo.ContracttBillInstallmentsDone.InstallNo,"
'MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.VATPayed, dbo.ContracttBillInstallmentsDone.VATValue, dbo.ContracttBillInstallmentsDone.ActVAT,"
'MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.Commisionvalue , dbo.ContracttBillInstallmentsDone.OldValuePayed, dbo.ContracttBillInstallmentsDone.PaymentType"
MySQL = MySQL & " FROM        dbo.Notes  "
'MySQL = MySQL & " dbo.ContracttBillInstallmentsDone RIGHT OUTER JOIN"
'MySQL = MySQL & "                       dbo.Notes ON dbo.ContracttBillInstallmentsDone.NoteID = dbo.Notes.NoteID "
MySQL = MySQL & "                       LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblNotesSales LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID ON dbo.Notes.NoteID = dbo.TblNotesSales.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_1 ON dbo.Notes.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_1 ON dbo.Notes.akarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqarDetai LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_2 ON dbo.TblAqarDetai.unittype = TblAkarUnit_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_2 ON dbo.TblAqarDetai.Aqarid = TblAqar_2.Aqarid ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
MySQL = MySQL & "                                     LEFT OUTER JOIN dbo.TblUsers AS tu"
MySQL = MySQL & "                                   ON  dbo.Notes.UserID = tu.UserID"
'Where (dbo.Notes.NoteID = 4441)
MySQL = MySQL & " Where "
MySQL = MySQL & " (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "        and CashingType >= 7"
'MySQL = MySQL & "        AND ISNULL(contNo, 0) <> 0"

If Me.DcbIqara2.BoundText <> "" Then
    StrWhere = StrWhere & " AND   dbo.Notes.akarid = " & val(Me.DcbIqara2.BoundText)
End If
    If GetTransIds <> "0" And Rd(3).value = False Then
    StrWhere = StrWhere & " and  dbo.Notes.NoteType in (" & GetTransIds & ") "
    End If


If Me.DcbUnitType2.BoundText <> "" Then
    StrWhere = StrWhere & " AND   dbo.Notes.unittype = " & val(Me.DcbUnitType2.BoundText)

End If
If Me.DcbUnitNo2.BoundText <> "" Then
    StrWhere = StrWhere & " AND   dbo.Notes.UnitNo = " & val(Me.DcbUnitNo2.BoundText)

End If



'''''''''
 If Not IsNull(Me.DtpDateFrom2.value) Then
                   StrWhere = StrWhere & " AND dbo.Notes.NoteDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo2.value) Then
                   StrWhere = StrWhere & " AND dbo.Notes.NoteDate<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
      End If
      
'If Me.DcbBranch2.BoundText <> "" Then
'    StrWhere = StrWhere & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch2.BoundText)
'End If



    MySQL = MySQL & StrWhere
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        MySQL = " SELECT     '' NoteID,'' UserName, '' NoteDate, '' NoteType,'' NoteSerial, '' NoteSerial1, 0 Note_Value, '' NoteDateH,"
        MySQL = MySQL & "                       '' ContractNo, '' ContNo, '' commission, '' rent, '' Water, '' FilterID, '' FIlterTotal, '' Instrunce,"
        MySQL = MySQL & "                       '' comX, '' ComY, '' CommissionOut, '' NoteOrBonID, '' comXold, "
        MySQL = MySQL & "                       '' NoteOrBonSereal, '' Telephone, '' CashingType, '' CusID, '' CusName, '' CusNamee"
        
        
       
         
        
       
      
     
    
    
       ' Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
       'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Exit Sub
    End If
'    Else
 'rs.MoveFirst
 
 

   Dim Sql2 As String, sql3 As String, s As String
    
    
    s = " select notes_all.NoteSerial1 IDFItWaiv,(IsNull(TblExpensesDet.Value,0))  Total,TblExpensesDet.des Remark,"
   's =s & "  NoteSerial IDFItWaiv,general_des nameDet,notes_all.NoteDate,
s = s & "     TblAqar.aqarname "
s = s & "                 From notes_all"
s = s & "      LEFT OUTER JOIN TblExpensesDet"
s = s & "      ON TblExpensesDet.ExpID= notes_all.NoteID "

s = s & "                 LEFT OUTER JOIN TblAqar ON TblExpensesDet.IqarID = TblAqar.Aqarid"
s = s & "                 Where NoteType = 3"
    
s = s & "                 and not (ToPriodDateH is null)"



If Not IsNull(Me.DtpDateFrom2.value) Then
        s = s & "                AND notes_all.NoteDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
        s = s & "                 And notes_all.NoteDate <= " & SQLDate(Me.DtpDateTo2.value, True) & ""
End If
If Me.DcbIqara2.BoundText <> "" Then
    s = s & " AND   TblAqar.Aqarid = " & val(Me.DcbIqara2.BoundText)
End If

's = s & "   GROUP BY aqarname,NoteSerial1"



    

Sql2 = s


s = " SELECT   "
s = s & "                          TotalExp = "
s = s & "                          ( SELECT Sum(ISNULL(Note_Value, 0))"
s = s & "                                     From notes_all"



s = s & "                 LEFT OUTER JOIN TblAqar ON notes_all.IqarID2 = TblAqar.Aqarid"
s = s & "                 Where NoteType = 3 "
s = s & "                 and not (ToPriodDateH is null)"
    




If Not IsNull(Me.DtpDateFrom2.value) Then
        s = s & "                AND notes_all.NoteDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
        s = s & "                 And notes_all.NoteDate <= " & SQLDate(Me.DtpDateTo2.value, True) & ""
End If
If Me.DcbIqara2.BoundText <> "" Then
    s = s & " AND   TblAqar.Aqarid = " & val(Me.DcbIqara2.BoundText)
End If


s = s & "       ),"

s = s & "       CountContract = (SELECT COUNT(*) FROM TblContract Where 1 = 1  "

If Not IsNull(Me.DtpDateFrom2.value) Then
        s = s & "                AND TblContract.ContDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
        s = s & "                 And TblContract.ContDate <= " & SQLDate(Me.DtpDateTo2.value, True) & ""
End If
If Me.DcbIqara2.BoundText <> "" Then
    s = s & " AND   dbo.TblContract.Iqar = " & val(Me.DcbIqara2.BoundText)
End If

If Me.DcbUnitType2.BoundText <> "" Then
    s = s & " AND   dbo.TblContract.unittype = " & val(Me.DcbUnitType2.BoundText)

End If
If Me.DcbUnitNo2.BoundText <> "" Then
    s = s & " AND   dbo.TblContract.UnitNo = " & val(Me.DcbUnitNo2.BoundText)

End If

s = s & "       ), "
s = s & "                          Cashing = SUM("
s = s & "                   Case Notes.NoteCashingType"
s = s & "                        WHEN 0 THEN (Note_Value2) + IsNull(Notes.Vat,0)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               ),"


s = s & "               Commission = (SELECT SUM("
s = s & "                          ISNULL(T2.CommissionsPayed, 0)"
s = s & "                      ) "
s = s & "               From dbo.TblContractInstallments"
s = s & "                      LEFT OUTER JOIN ContracttBillInstallmentsDone T2"
s = s & "                           ON  T2.istallid = TblContractInstallments.ID"
s = s & "               Where IsNull(TblContractInstallments.ContNo,0) <>0"

        If Not IsNull(Me.DtpDateFrom2.value) Then
                   s = s & "                AND T2.RecordDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
        End If
        If Not IsNull(Me.DtpDateTo2.value) Then
                   s = s & "                 And T2.RecordDate <= " & SQLDate(Me.DtpDateTo2.value, True) & ""
        End If
s = s & "               ),"


s = s & "               TelandNetPayed = (SELECT SUM("
s = s & "                          ISNULL(T2.TelandNetPayed, 0)"
s = s & "                      ) "
s = s & "               From dbo.TblContractInstallments"
s = s & "                      LEFT OUTER JOIN ContracttBillInstallmentsDone T2"
s = s & "                           ON  T2.istallid = TblContractInstallments.ID"
s = s & "               Where IsNull(TblContractInstallments.ContNo,0) <>0"

If Not IsNull(Me.DtpDateFrom2.value) Then
           s = s & "                AND T2.RecordDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
           s = s & "                 And T2.RecordDate <= " & SQLDate(Me.DtpDateTo2.value, True) & ""
End If
s = s & "               )"

s = s & "               ,"
s = s & "               Arbon             = SUM(CASE Notes.CashingType WHEN 9 THEN (Note_Value) ELSE 0 END),"
s = s & "               ValueTransfer     = SUM("
s = s & "                   Case Notes.NoteCashingType"
s = s & "                        WHEN 2 THEN (Note_Value2)  + IsNull(Notes.Vat,0)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               )"
s = s & "        From Notes Where 1 = 1 and NoteType = 4 and CashingType >= 7"
s = s & StrWhere
sql3 = s


 print_report MySQL, Sql2, sql3
   ' End If

End Sub

Function print_report(Optional NoteSerial As String, Optional SQLStr As String = "", Optional sqlStr2 As String = "")
     
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
   
If Me.RdTotal.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllIncomAndExpenRep.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllIncomAndExpenRep.rpt"
            
       End If
End If

If Rd(4).value = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFilterWiaverFinancial.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFilterWiaverFinancial.rpt"
       End If
End If


If Rd(3).value = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportExpensesStyle1Iqar.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportExpensesStyle1Iqar.rpt"
       End If
End If
If Me.RdAnalis.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysIncomAndExpenRep.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysIncomAndExpenRep.rpt"
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

    If SQLStr <> "" Then
        Dim RsData2  As New ADODB.Recordset
        Dim RsData3  As New ADODB.Recordset
         
        RsData2.Open SQLStr, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("Expens").Database.SetDataSource RsData2
        
        RsData3.Open sqlStr2, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("TotalIq").Database.SetDataSource RsData3
        
    End If

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value
    xReport.ParameterFields(9).AddCurrentValue DtpDateFromH.value
    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
    xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If

  Dim total As String
  Dim totl As Double


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



 
Private Sub TxtSearch2_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch2.Text, EmpID
        DcbIqara2.BoundText = EmpID
        DcbIqara2_Click (0)
    End If
End Sub
