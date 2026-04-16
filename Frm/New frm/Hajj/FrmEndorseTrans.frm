VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEndorseTrans 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«⁄ „«œ ≈—ﬂ«» «·ÕÃ«Ã"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "FrmEndorseTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10785
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
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10785
      _cx             =   19024
      _cy             =   9975
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
      Caption         =   "»Ì«‰«  «”«”Ì…| Ê“Ì⁄ «·Õ«›·« "
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5235
         Left            =   45
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Width           =   10695
         _cx             =   18865
         _cy             =   9234
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
         Begin VB.TextBox ApproveID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1920
            Width           =   3810
         End
         Begin VB.TextBox TxtSession 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   2280
            Width           =   3810
         End
         Begin VB.TextBox TotOlds 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7755
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox TotYoungs 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7755
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   720
            Width           =   1410
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   7755
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox TXtLargPrice 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   5835
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox TxtSmalPrice 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   5835
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   720
            Width           =   1410
         End
         Begin VB.TextBox TxtTotalPrice 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5835
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1080
            Width           =   1410
         End
         Begin VB.TextBox TxtGroupName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1560
            Width           =   3810
         End
         Begin VB.TextBox TxtPassword 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   5835
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2160
            Width           =   3330
         End
         Begin VB.TextBox TxtPArCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5835
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1800
            Width           =   3330
         End
         Begin VB.TextBox Remark 
            Alignment       =   1  'Right Justify
            Height          =   690
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   2880
            Width           =   8925
         End
         Begin MSDataListLib.DataCombo SeasonsID 
            Height          =   315
            Left            =   240
            TabIndex        =   42
            Top             =   840
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPath 
            Height          =   315
            Left            =   240
            TabIndex        =   43
            Top             =   120
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID 
            Height          =   315
            Left            =   240
            TabIndex        =   44
            Top             =   1200
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Nationality 
            Height          =   315
            Left            =   5835
            TabIndex        =   45
            Top             =   1440
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   1215
            Left            =   120
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   3600
            Width           =   10425
            _cx             =   18389
            _cy             =   2143
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
            Begin VB.TextBox TxtPhone 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   720
               Width           =   1650
            End
            Begin VB.TextBox txtReceptID 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   360
               Width           =   1770
            End
            Begin VB.TextBox TxtReceptOffice 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   360
               Width           =   1650
            End
            Begin VB.TextBox TxtReceptName 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   360
               Width           =   3690
            End
            Begin MSDataListLib.DataCombo DcbLocation 
               Height          =   315
               Left            =   5400
               TabIndex        =   74
               Top             =   720
               Width           =   3690
               _ExtentX        =   6509
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker ReceptTime 
               Height          =   288
               Left            =   120
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   720
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   94568450
               CurrentDate     =   37140
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Êﬁ "
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1785
               TabIndex        =   83
               Top             =   720
               Width           =   750
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Â« ›"
               Height          =   330
               Index           =   25
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   720
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Êﬁ⁄ «· Õ„Ì·"
               Height          =   180
               Index           =   23
               Left            =   9270
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÂÊÌ…"
               Height          =   330
               Index           =   22
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬂ »"
               Height          =   330
               Index           =   21
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” ·„"
               Height          =   330
               Index           =   20
               Left            =   9030
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„” ·„"
               Height          =   330
               Index           =   19
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   0
               Width           =   1185
            End
         End
         Begin MSDataListLib.DataCombo DcbVehicleType 
            Height          =   315
            Left            =   240
            TabIndex        =   70
            Top             =   480
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDepandID 
            Height          =   315
            Left            =   5835
            TabIndex        =   90
            Top             =   2520
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«⁄ „«œ"
            Height          =   285
            Index           =   29
            Left            =   9420
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   2520
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›∆…"
            Height          =   300
            Index           =   6
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   480
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   180
            Index           =   10
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   840
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·œÊ—…"
            Height          =   300
            Index           =   12
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   2280
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«⁄ „«œ"
            Height          =   300
            Index           =   13
            Left            =   4230
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1920
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œÿ «·”Ì—"
            Height          =   300
            Index           =   14
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ»«—"
            Height          =   285
            Index           =   9
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   360
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«‰’«›"
            Height          =   165
            Index           =   5
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   720
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "««·„Ã„Ê⁄"
            Height          =   285
            Index           =   1
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄œœ"
            Height          =   405
            Index           =   15
            Left            =   7695
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   120
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   405
            Index           =   16
            Left            =   6135
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ƒ””…"
            Height          =   300
            Index           =   7
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ã„Ê⁄…"
            Height          =   300
            Index           =   17
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ã‰”Ì…"
            Height          =   285
            Index           =   0
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1440
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·”—Ì"
            Height          =   300
            Index           =   11
            Left            =   9645
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   2160
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·»«—ﬂÊœ"
            Height          =   285
            Index           =   18
            Left            =   9615
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1800
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   405
            Index           =   3
            Left            =   9495
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3000
            Width           =   930
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5235
         Left            =   11430
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   45
         Width           =   10695
         _cx             =   18865
         _cy             =   9234
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
         Begin VB.TextBox TxtCapacity 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   4800
            Width           =   1770
         End
         Begin VB.TextBox TxtNoVehicle 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   3255
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   4800
            Width           =   1770
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   4515
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Width           =   10485
            _cx             =   18494
            _cy             =   7964
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
            Rows            =   12
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmEndorseTrans.frx":038A
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
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   270
            Left            =   8640
            TabIndex        =   84
            ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
            Top             =   4800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
            ButtonImage     =   "FrmEndorseTrans.frx":04F9
            ButtonImageDisabled=   "FrmEndorseTrans.frx":6D5B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   270
            Left            =   6840
            TabIndex        =   85
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   4800
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ· "
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
            ButtonImage     =   "FrmEndorseTrans.frx":25F45
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·”⁄…"
            Height          =   330
            Index           =   27
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   4800
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Õ«›·« "
            Height          =   330
            Index           =   26
            Left            =   4935
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   4800
            Width           =   1320
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic7 
      Height          =   750
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7680
      Width           =   10725
      _cx             =   18918
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   510
         Index           =   0
         Left            =   9840
         TabIndex        =   2
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":2C7A7
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
         Height          =   510
         Index           =   1
         Left            =   8940
         TabIndex        =   3
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":33009
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
         Height          =   510
         Index           =   2
         Left            =   8055
         TabIndex        =   4
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":3986B
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
         Height          =   510
         Index           =   3
         Left            =   7170
         TabIndex        =   5
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":400CD
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
         Height          =   510
         Index           =   4
         Left            =   6240
         TabIndex        =   6
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":4692F
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
         Height          =   510
         Index           =   6
         Left            =   1050
         TabIndex        =   7
         Top             =   120
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":4D191
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
         Height          =   510
         Left            =   105
         TabIndex        =   8
         Top             =   120
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":76DB3
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
         Height          =   510
         Index           =   7
         Left            =   5370
         TabIndex        =   9
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
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
         ButtonImage     =   "FrmEndorseTrans.frx":7D615
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
         Height          =   510
         Index           =   9
         Left            =   1875
         TabIndex        =   10
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   900
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
         ButtonImage     =   "FrmEndorseTrans.frx":83E77
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
         Height          =   510
         Index           =   5
         Left            =   4200
         TabIndex        =   86
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… ”‰œ «” ·«„"
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
         ButtonImage     =   "FrmEndorseTrans.frx":8A6D9
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
         Height          =   510
         Index           =   8
         Left            =   2760
         TabIndex        =   87
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   900
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… »Ê«ﬁÌ"
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
         ButtonImage     =   "FrmEndorseTrans.frx":90F3B
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   735
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   10725
      _cx             =   18918
      _cy             =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   22.5
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
      Caption         =   "«⁄ „«œ ≈—ﬂ«» «·ÕÃ«Ã"
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
      CaptionStyle    =   1
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   13
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmEndorseTrans.frx":9779D
         ColorButton     =   -2147483634
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
         TabIndex        =   14
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmEndorseTrans.frx":97B37
         ColorButton     =   -2147483634
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
         TabIndex        =   15
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmEndorseTrans.frx":97ED1
         ColorButton     =   -2147483634
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
         TabIndex        =   16
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmEndorseTrans.frx":9826B
         ColorButton     =   -2147483634
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic6 
      Height          =   510
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7080
      Width           =   10800
      _cx             =   19050
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
      Begin MSDataListLib.DataCombo DcbUser 
         Height          =   315
         Left            =   6120
         TabIndex        =   69
         Top             =   120
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ê«”ÿ…"
         Height          =   180
         Index           =   28
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   120
         Width           =   840
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   330
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   330
         Index           =   2
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   330
         Index           =   4
         Left            =   870
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   1020
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   615
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   720
      Width           =   10755
      _cx             =   18971
      _cy             =   1085
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
      Begin VB.TextBox ID 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   8160
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   120
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker SDate 
         Height          =   300
         Left            =   6120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   120
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94568451
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo BranchID 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Dynamic_Byte.NourHijriCal RecordDateH 
         Height          =   315
         Left            =   4440
         TabIndex        =   88
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   420
         Index           =   24
         Left            =   3630
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7185
         TabIndex        =   26
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”·”·"
         Height          =   330
         Index           =   8
         Left            =   9420
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FrmEndorseTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Private Sub ApproveID_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.ApproveID.Text, 1)
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
sql = "SELECT     dbo.TblEndorseTransDet.ID, dbo.TblEndorseTransDet.EnTransID, dbo.TblEndorseTransDet.EmpID,dbo.TblEndorseTransDet.Capacity, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
sql = sql & "                       dbo.TblEmployee.Emp_Namee, dbo.TblEndorseTransDet.CarID, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
sql = sql & "                       dbo.TblCarsData.BoardNO , dbo.TblCarsData.Model, dbo.TblCarsData.OperatorN"
sql = sql & "  FROM         dbo.TBLCarTypes RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblCarsData ON dbo.TBLCarTypes.id = dbo.TblCarsData.CarsTypeId RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblEndorseTransDet ON dbo.TblCarsData.id = dbo.TblEndorseTransDet.CarID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee ON dbo.TblEndorseTransDet.EmpID = dbo.TblEmployee.Emp_ID"
sql = sql & "  Where (dbo.TblEndorseTransDet.EnTransID =" & val(ID.Text) & ") "


  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Capacity")) = IIf(IsNull(Rs1("Capacity").value), 0, Rs1("Capacity").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(Rs1("BoardNO").value), "", Rs1("BoardNO").value)
                   .TextMatrix(i, .ColIndex("ModelID")) = IIf(IsNull(Rs1("Model").value), "", Rs1("Model").value)
                   .TextMatrix(i, .ColIndex("CarNo")) = IIf(IsNull(Rs1("OperatorN").value), "", Rs1("OperatorN").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), "", Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(Rs1("CarID").value), 0, Rs1("CarID").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
    Function ChekDepn() As Boolean
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblEndorseTrans where id=" & val(ID.Text) & " and FlagDepand =1  "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
ChekDepn = True
Else
ChekDepn = False
End If
End Function
Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Select Case Index
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
             BranchID.BoundText = Current_branch
            ID.Text = CStr(new_id("TblEndorseTrans", "ID", "", True))
         SeasonsID.BoundText = GetMosim(1)
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
              If ChekDepn() = True Then
           MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–« «·«⁄ „«œ „— »ÿ »‘«‘… «·„ÿ«·»« "
           Exit Sub
           End If
GridInstallments.Rows = GridInstallments.Rows + 1
            TxtModFlg.Text = "E"
  
        Case 2
If val(SeasonsID.BoundText) = 0 Or SeasonsID.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„Ê”„"
Else
MsgBox "Please Select Select Current Year"
End If
SeasonsID.SetFocus
Exit Sub
End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
        If ChekDepn() = True Then
           MsgBox "«Ì„ﬂ‰ «·Õ–› Â–« «·«⁄ „«œ „— »ÿ »‘«‘…  ›«’Ì· «·«⁄ „«œ"
           Exit Sub
           End If
            Del_Action

        Case 5
          print_report2 , 1
        Case 6
                Unload Me
         Case 7
                print_report2 , 0
       Case 8
                print_report2 , 2
         Case 9
         Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "EndorseTrans"
         FrmSearch_Hajj.show
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub




Private Sub CmdAttach_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments ID.Text, "20911201603"
End Sub

Private Sub Form_Activate()
'    txtid.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
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

Private Sub Fill_Combos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
  
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches BranchID
   Dcombos.GetTblShrines DcbPath
   Dcombos.GETNationality Nationality
   Dcombos.GetTypeDependence Me.DcbDepandID
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
 End If
 str = str & " where Omra_Hajj=1"
   fill_combo SeasonsID, str
    If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblvehicleType  "
   Else
   str = " select id , nameE from TblvehicleType  "
 End If
fill_combo DcbVehicleType, str
   Dcombos.GetTblLocations Me.DcbLocation
 If SystemOptions.UserInterface = ArabicInterface Then
    str = " select  ID , Name   from TblTourismCompanies "
Else
    str = " Select ID , NameE from  TblTourismCompanies "
End If
fill_combo CompanyID, str
  ' str = " select id , name from tblcompaniesgroup "
  ' fill_combo GroupID, str
   
End Sub


Private Sub Form_Load()
 '   On Error GoTo ErrTrap
        Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  „·› «⁄ „«œ «·ÕÃ  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 '   Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset

    Dim StrSQL As String
    StrSQL = ""
    
     If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = "SELECT  *  From TblEndorseTrans    "
    Else
            StrSQL = "SELECT  *  From TblEndorseTrans"
    End If
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub

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
       
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
   
   ' lbl(7).Caption = " Name En"
   ' lbl(3).Caption = " Name Ar"
   ' lbl(8).Caption = "Process No"
   ' lbl(0).Caption = "Minister No."
   ' Label3.Caption = "School Manager"
   ' Label1.Caption = "Managerial Area"
   ' Label2.Caption = "City"
   '' lbl(5).Caption = "Student Count"
   ' lbl(6).Caption = "Custom"
   ' lbl(1).Caption = "School Type"
   ' lbl(10).Caption = "Telephone"
   ' Label4.Caption = "Supervisor Code"
   ' Label6.Caption = "Supervisor"
   ' Label5.Caption = "Student Gender"
   '' Me.Caption = "School Data"
   ' EleHeader.Caption = Me.Caption
   '
   ' lbl(2).Caption = "Current Record"
   ' lbl(4).Caption = "NO. Recordes"
'
'    Me.Cmd(0).Caption = "New"
'    Me.Cmd(1).Caption = "Edit"
''    Me.Cmd(2).Caption = "Save"
'    Me.Cmd(3).Caption = "Undo"
'    Me.Cmd(4).Caption = "Delete"
'    'Me.Cmd(5).Caption = "Search"
'    Me.Cmd(6).Caption = "Exit"
''    Me.Cmd(7).Caption = "Print"
'   CmdAttach.Caption = "Attachment"

'lbl(9).Caption = "Last Contract"



End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã    "
    LogTexte = " Exit Window " & "  "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub




Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
     Dim Capacity As Integer
    IntCounter = 0
    Capacity = 0
    With GridInstallments

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("CarID"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            Capacity = Capacity + val(.TextMatrix(i, .ColIndex("Capacity")))
            End If
        Next i
 
    End With
    TxtNoVehicle.Text = IntCounter
      TxtCapacity.Text = Capacity

End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Rs4 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim StrComboList As String
    With GridInstallments
    
        Select Case .ColKey(Col)
    Case "EmpName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID"), False, True)
                .TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
                StrSQL = "Select Fullcode From TblEmployee where Emp_ID =" & val(.TextMatrix(Row, .ColIndex("EmpID"))) & ""
                Set Rs4 = New ADODB.Recordset
                
                StrSQL = "  SELECT     dbo.TblCarsData.id, dbo.TblCarsData.Model, dbo.TblCarsData.Capacity, dbo.TblCarsData.BoardNO, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
               StrSQL = StrSQL & "        dbo.TblCarsData.Emp_id , dbo.TblEmployee.emp_name, dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee"
               StrSQL = StrSQL & "   ,OperatorN      FROM         dbo.TblCarsData LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id"
               StrSQL = StrSQL & "     Where (dbo.TblCarsData.Emp_id =" & val(.TextMatrix(Row, .ColIndex("EmpID"))) & ")"
              Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                  If Rs4.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CarID")) = IIf(IsNull(Rs4("id").value), 0, Rs4("id").value)
                .TextMatrix(Row, .ColIndex("CarNo")) = IIf(IsNull(Rs4("OperatorN").value), 0, Rs4("OperatorN").value)
                .TextMatrix(Row, .ColIndex("BoardNO")) = IIf(IsNull(Rs4("BoardNO").value), 0, Rs4("BoardNO").value)
                .TextMatrix(Row, .ColIndex("ModelID")) = IIf(IsNull(Rs4("Model").value), 0, Rs4("Model").value)
                .TextMatrix(Row, .ColIndex("Fullcode")) = IIf(IsNull(Rs4("Fullcode").value), "", Rs4("Fullcode").value)
                .TextMatrix(Row, .ColIndex("Capacity")) = IIf(IsNull(Rs4("Capacity").value), 0, Rs4("Capacity").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(Rs4("name").value), "", Rs4("name").value)
                
                Else
                .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(Rs4("namee").value), "", Rs4("namee").value)
             '   .TextMatrix(Row, .ColIndex("EmpName")) = IIf(IsNull(Rs4("Emp_Namee").value), "", Rs4("Emp_Namee").value)
                End If
                Else
             .TextMatrix(Row, .ColIndex("CarID")) = ""
             .TextMatrix(Row, .ColIndex("BoardNO")) = ""
             .TextMatrix(Row, .ColIndex("CarNo")) = ""
             .TextMatrix(Row, .ColIndex("name")) = ""
             .TextMatrix(Row, .ColIndex("ModelID")) = ""
             .TextMatrix(Row, .ColIndex("Fullcode")) = ""
             .TextMatrix(Row, .ColIndex("Capacity")) = ""
                End If
         Case "Fullcode"
                
             StrSQL = "  SELECT     dbo.TblCarsData.id,dbo.TblCarsData.OperatorN,dbo.TblCarsData.Capacity, dbo.TblCarsData.Model, dbo.TblCarsData.BoardNO, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
               StrSQL = StrSQL & "        dbo.TblCarsData.Emp_id , dbo.TblEmployee.emp_name, dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee"
               StrSQL = StrSQL & "         FROM         dbo.TblCarsData LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id"
               StrSQL = StrSQL & "     Where (dbo.TblEmployee.fullcode ='" & (.TextMatrix(Row, .ColIndex("Fullcode"))) & "')"
                Set Rs4 = New ADODB.Recordset
                Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs4.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CarID")) = IIf(IsNull(Rs4("id").value), 0, Rs4("id").value)
                .TextMatrix(Row, .ColIndex("CarNo")) = IIf(IsNull(Rs4("OperatorN").value), 0, Rs4("OperatorN").value)
                .TextMatrix(Row, .ColIndex("BoardNO")) = IIf(IsNull(Rs4("BoardNO").value), 0, Rs4("BoardNO").value)
                .TextMatrix(Row, .ColIndex("ModelID")) = IIf(IsNull(Rs4("Model").value), 0, Rs4("Model").value)
                .TextMatrix(Row, .ColIndex("EmpID")) = IIf(IsNull(Rs4("Emp_id").value), 0, Rs4("Emp_id").value)
                .TextMatrix(Row, .ColIndex("Capacity")) = IIf(IsNull(Rs4("Capacity").value), 0, Rs4("Capacity").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(Rs4("name").value), "", Rs4("name").value)
                .TextMatrix(Row, .ColIndex("EmpName")) = IIf(IsNull(Rs4("Emp_Name").value), "", Rs4("Emp_Name").value)
                Else
                .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(Rs4("namee").value), "", Rs4("namee").value)
                .TextMatrix(Row, .ColIndex("EmpName")) = IIf(IsNull(Rs4("Emp_Namee").value), "", Rs4("Emp_Namee").value)
                End If
                Else
                .TextMatrix(Row, .ColIndex("EmpName")) = ""
                .TextMatrix(Row, .ColIndex("EmpID")) = 0
                .TextMatrix(Row, .ColIndex("CarID")) = ""
                .TextMatrix(Row, .ColIndex("BoardNO")) = ""
                .TextMatrix(Row, .ColIndex("CarNo")) = ""
                .TextMatrix(Row, .ColIndex("name")) = ""
                .TextMatrix(Row, .ColIndex("ModelID")) = ""
                .TextMatrix(Row, .ColIndex("Capacity")) = 0
                End If
    Case "CarNo"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CarID"), False, True)
                .TextMatrix(Row, .ColIndex("CarID")) = StrAccountCode
               StrSQL = "  SELECT     dbo.TblCarsData.id, dbo.TblCarsData.Model, dbo.TblCarsData.Capacity , dbo.TblCarsData.BoardNO, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
               StrSQL = StrSQL & "        dbo.TblCarsData.Emp_id , dbo.TblEmployee.emp_name, dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee"
               StrSQL = StrSQL & "         FROM         dbo.TblCarsData LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id"
               StrSQL = StrSQL & "     Where (dbo.TblCarsData.ID =" & val(.TextMatrix(Row, .ColIndex("CarID"))) & ")"
                 Set Rs4 = New ADODB.Recordset
                Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs4.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("Fullcode")) = IIf(IsNull(Rs4("Fullcode").value), "", Rs4("Fullcode").value)
                .TextMatrix(Row, .ColIndex("BoardNO")) = IIf(IsNull(Rs4("BoardNO").value), 0, Rs4("BoardNO").value)
                .TextMatrix(Row, .ColIndex("ModelID")) = IIf(IsNull(Rs4("Model").value), 0, Rs4("Model").value)
                .TextMatrix(Row, .ColIndex("EmpID")) = IIf(IsNull(Rs4("Emp_id").value), 0, Rs4("Emp_id").value)
                .TextMatrix(Row, .ColIndex("Capacity")) = IIf(IsNull(Rs4("Capacity").value), 0, Rs4("Capacity").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(Rs4("name").value), "", Rs4("name").value)
                .TextMatrix(Row, .ColIndex("EmpName")) = IIf(IsNull(Rs4("Emp_Name").value), "", Rs4("Emp_Name").value)
                Else
                .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(Rs4("namee").value), "", Rs4("namee").value)
                .TextMatrix(Row, .ColIndex("EmpName")) = IIf(IsNull(Rs4("Emp_Namee").value), "", Rs4("Emp_Namee").value)
                End If
                Else
                .TextMatrix(Row, .ColIndex("EmpName")) = ""
                .TextMatrix(Row, .ColIndex("EmpID")) = 0
                .TextMatrix(Row, .ColIndex("Fullcode")) = ""
                .TextMatrix(Row, .ColIndex("name")) = ""
                .TextMatrix(Row, .ColIndex("ModelID")) = ""
                .TextMatrix(Row, .ColIndex("Capacity")) = ""
                End If

                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If
    End With

    ReLineGrid
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "name"
Cancel = True
Case "ModelID"
Cancel = True
Case "BoardNO"
Cancel = True
Case "Fullcode"
.ComboList = ""
Case "Capacity"
.ComboList = ""
End Select
End With
End Sub

Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim Rs4 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrComboList As String
    Dim ClsAcc As New ClsAccounts
    With GridInstallments
Set Rs4 = New ADODB.Recordset
      Select Case .ColKey(Col)
  Case "CarNo"
                StrSQL = " select ID,OperatorN from TblCarsData"
                StrSQL = StrSQL & " where (NOT (OperatorN IS NULL)) "
                Rs4.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = .BuildComboList(Rs4, "OperatorN", "ID")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
 Case "EmpName"
      
                StrSQL = "  select   e.Emp_ID Emp_ID , e.Emp_Name,e.Emp_NameE   from TblEmployee e, TblEmpJobsTypes  j"
                StrSQL = StrSQL & "   Where e.JobTypeID = j.JobTypeID"
                StrSQL = StrSQL & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
                Rs4.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(Rs4, "Emp_Name", "Emp_ID")
                  Else
                  StrComboList = .BuildComboList(Rs4, "Emp_NameE", "Emp_ID")
                  End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
        End Select

    End With
End Sub

Private Sub ID_Change()
ApproveID.Text = ID.Text
End Sub

Private Sub ISButton3_Click()
On Error Resume Next
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub

Private Sub ISButton4_Click()
 On Error Resume Next
 Me.GridInstallments.Clear flexClearScrollable, flexClearEverything
 Me.GridInstallments.Rows = 1
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    SDate.value = ToGregorianDate(RecorddateH.value)
    End If
End Sub

Private Sub SDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecorddateH.value = ToHijriDate(SDate.value)
End If
End Sub

Private Sub TotOlds_Change()
If Me.TxtModFlg.Text <> "R" Then
txtTotalPrice.Text = val(TXtLargPrice.Text) * val(TotOlds.Text) + val(TxtSmalPrice.Text) * val(TotYoungs.Text)
Total.Text = val(TotYoungs.Text) + val(TotOlds.Text)
End If
End Sub

Private Sub TotOlds_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TotOlds.Text, 1)
End Sub

Private Sub TotYoungs_Change()
If Me.TxtModFlg.Text <> "R" Then
txtTotalPrice.Text = val(TXtLargPrice.Text) * val(TotOlds.Text) + val(TxtSmalPrice.Text) * val(TotYoungs.Text)
Total.Text = val(TotYoungs.Text) + val(TotOlds.Text)
End If
End Sub

Private Sub TotYoungs_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TotYoungs.Text, 1)
End Sub

Private Sub TXtLargPrice_Change()
If Me.TxtModFlg.Text <> "R" Then
txtTotalPrice.Text = val(TXtLargPrice.Text) * val(TotOlds.Text) + val(TxtSmalPrice.Text) * val(TotYoungs.Text)
Total.Text = val(TotYoungs.Text) + val(TotOlds.Text)
End If
End Sub

Private Sub TXtLargPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TXtLargPrice.Text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «⁄ „«œ «—ﬂ«» ÕÃ«Ã"
            Else
                Me.Caption = "School  Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(9).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            ID.locked = True
      

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            'pnlHeader.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
           ' pnlHeader.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã «·ÕÃ“ (  ⁄œÌ· )"
            Else
                Me.Caption = "Booking Request Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            ID.locked = True
          ' pnlHeader.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String, Optional Index As Integer = 0)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  If Index = 2 Then
MySQL = "   SELECT     dbo.TblEndorseTrans.SDate, dbo.TblEndorseTrans.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "                       dbo.TblEndorseTrans.ApproveID, dbo.TblEndorseTrans.Nationality, dbo.Nationality.name, dbo.Nationality.namee, dbo.TblEndorseTrans.TotOlds,"
MySQL = MySQL & "                       dbo.TblEndorseTrans.TotYoungs, dbo.TblEndorseTrans.Total, dbo.TblEndorseTrans.Remark, dbo.TblEndorseTrans.GroupName, dbo.TblEndorseTrans.ReceptName,"
MySQL = MySQL & "                       dbo.TblEndorseTrans.ReceptID, dbo.TblEndorseTrans.ReceptOffice, dbo.TblEndorseTrans.PArCode, dbo.TblEndorseTrans.Password, dbo.TblEndorseTrans.[Session],"
MySQL = MySQL & "                       dbo.TblEndorseTrans.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE, dbo.TblEndorseTrans.LargPrice,"
MySQL = MySQL & "                       dbo.TblEndorseTrans.SmalPrice, dbo.TblEndorseTrans.TotalPrice, dbo.TblEndorseTrans.ReceptTime, dbo.TblEndorseTrans.Phone, dbo.TblEndorseTrans.NoVehicle,"
MySQL = MySQL & "                       dbo.TblEndorseTrans.Capacity, dbo.TblEndorseTrans.LocationID, dbo.TblLocations.Name AS LocationName, dbo.TblLocations.NameE AS LocationNameE,"
MySQL = MySQL & "                       dbo.TblEndorseTrans.VehicleType, dbo.TblVehicleType.Name AS VehicleName, dbo.TblVehicleType.NameE AS VehicleNameE, dbo.TblEndorseTrans.ID,"
MySQL = MySQL & "                       dbo.TblEndorseTransDet.Capacity AS CapacityDet, dbo.TblEndorseTransDet.CarID, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name AS CarTypename,"
MySQL = MySQL & "                       dbo.TBLCarTypes.namee AS CarTypenameE, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.OperatorN, dbo.TblEndorseTransDet.EmpID,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEndorseTrans.CompanyID,"
MySQL = MySQL & "                       dbo.TblTourismCompanies.Name AS ComName, dbo.TblTourismCompanies.NameE AS ComNameE, dbo.TblCompaniesGroup.Name AS SeasonName,"
MySQL = MySQL & "                       dbo.TblCompaniesGroup.NameE AS SeasonNameE, dbo.TblEndorseTrans.SeasonsID, dbo.GetNoHajSaml(dbo.TblEndorseTrans.ID) AS NoSmal,"
MySQL = MySQL & "                       dbo.GetNoHajLarg(dbo.TblEndorseTrans.ID) AS NoLarg"
MySQL = MySQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEndorseTrans ON dbo.Nationality.id = dbo.TblEndorseTrans.Nationality LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblTourismCompanies ON dbo.TblEndorseTrans.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEndorseTransDet ON dbo.TblEmployee.Emp_ID = dbo.TblEndorseTransDet.EmpID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCarsData LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id ON dbo.TblEndorseTransDet.CarID = dbo.TblCarsData.id ON"
MySQL = MySQL & "                       dbo.TblEndorseTrans.ID = dbo.TblEndorseTransDet.EnTransID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblVehicleType ON dbo.TblEndorseTrans.VehicleType = dbo.TblVehicleType.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblLocations ON dbo.TblEndorseTrans.LocationID = dbo.TblLocations.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblShrines ON dbo.TblEndorseTrans.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblEndorseTrans.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCompaniesGroup ON dbo.TblEndorseTrans.SeasonsID = dbo.TblCompaniesGroup.ID"
MySQL = MySQL & "     Where dbo.TblEndorseTrans.ID =  " & val(ID.Text)
  Else
  
MySQL = " SELECT     dbo.TblEndorseTrans.SDate, dbo.TblEndorseTrans.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
MySQL = MySQL & "                      dbo.TblEndorseTrans.ApproveID, dbo.TblEndorseTrans.Nationality, dbo.Nationality.name, dbo.Nationality.namee, dbo.TblEndorseTrans.TotOlds,"
MySQL = MySQL & "                      dbo.TblEndorseTrans.TotYoungs, dbo.TblEndorseTrans.Total, dbo.TblEndorseTrans.Remark, dbo.TblEndorseTrans.GroupName, dbo.TblEndorseTrans.ReceptName,"
MySQL = MySQL & "                      dbo.TblEndorseTrans.ReceptID, dbo.TblEndorseTrans.ReceptOffice, dbo.TblEndorseTrans.PArCode, dbo.TblEndorseTrans.Password, dbo.TblEndorseTrans.[Session],"
MySQL = MySQL & "                      dbo.TblEndorseTrans.PathID, dbo.TblShrines.Name AS PATHName, dbo.TblShrines.NameE AS PATHNameE, dbo.TblEndorseTrans.LargPrice,"
MySQL = MySQL & "                      dbo.TblEndorseTrans.SmalPrice, dbo.TblEndorseTrans.TotalPrice, dbo.TblEndorseTrans.ReceptTime, dbo.TblEndorseTrans.Phone, dbo.TblEndorseTrans.NoVehicle,"
MySQL = MySQL & "                      dbo.TblEndorseTrans.Capacity, dbo.TblEndorseTrans.LocationID, dbo.TblLocations.Name AS LocationName, dbo.TblLocations.NameE AS LocationNamee,"
MySQL = MySQL & "                      dbo.TblEndorseTrans.VehicleType, dbo.TblVehicleType.Name AS VehicleName, dbo.TblVehicleType.NameE AS VehicleNameE, dbo.TblEndorseTrans.ID,"
MySQL = MySQL & "                      dbo.TblEndorseTransDet.Capacity AS CapacityDet, dbo.TblEndorseTransDet.CarID, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name AS CarTypename,"
MySQL = MySQL & "                      dbo.TBLCarTypes.namee AS CarTypenameE, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.OperatorN, dbo.TblEndorseTransDet.EmpID,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEndorseTrans.CompanyID,"
MySQL = MySQL & "                      dbo.TblTourismCompanies.Name AS ComName, dbo.TblTourismCompanies.NameE AS ComNameE, dbo.TblCompaniesGroup.Name AS SeasonName,"
MySQL = MySQL & "                      dbo.TblCompaniesGroup.NameE AS SeasonNameE, dbo.TblEndorseTrans.SeasonsID, dbo.TblEndorseTrans.DepandID, dbo.TblEndorseTrans.RecordDateH,"
MySQL = MySQL & "                      dbo.TblTypeDependence.Name AS DepandName, dbo.TblTypeDependence.NameE AS DepandNameE"
MySQL = MySQL & " FROM         dbo.TblTypeDependence RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEndorseTrans ON dbo.TblTypeDependence.ID = dbo.TblEndorseTrans.DepandID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Nationality ON dbo.TblEndorseTrans.Nationality = dbo.Nationality.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblTourismCompanies ON dbo.TblEndorseTrans.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEndorseTransDet ON dbo.TblEmployee.Emp_ID = dbo.TblEndorseTransDet.EmpID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarsData LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id ON dbo.TblEndorseTransDet.CarID = dbo.TblCarsData.id ON"
MySQL = MySQL & "                      dbo.TblEndorseTrans.ID = dbo.TblEndorseTransDet.EnTransID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVehicleType ON dbo.TblEndorseTrans.VehicleType = dbo.TblVehicleType.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblLocations ON dbo.TblEndorseTrans.LocationID = dbo.TblLocations.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblShrines ON dbo.TblEndorseTrans.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblEndorseTrans.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCompaniesGroup ON dbo.TblEndorseTrans.SeasonsID = dbo.TblCompaniesGroup.ID"
MySQL = MySQL & "   where dbo.TblEndorseTrans.ID =  " & val(ID.Text)
End If
If Index = 2 Then
     If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ReceptBusesReamin.rpt"
     Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ReceptBusesReamin.rpt"
     End If
ElseIf Index = 1 Then
     If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ReceptBuses.rpt"
     Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ReceptBuses.rpt"
     End If
  Else
      If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EndorseTrans.rpt"
     Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EndorseTrans.rpt"
    End If
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
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
Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
     If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
   
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    RecorddateH.value = IIf(IsNull(rs("RecordDateH").value), ToHijriDate(Date), rs("RecordDateH").value)
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    BranchID.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    DcbDepandID.BoundText = IIf(IsNull(rs("DepandID").value), "", Trim(rs("DepandID").value))
    SeasonsID.BoundText = IIf(IsNull(rs("SeasonsID").value), "", Trim(rs("SeasonsID").value))
    ApproveID.Text = IIf(IsNull(rs("ApproveID").value), "", Trim(rs("ApproveID").value))
    Nationality.BoundText = IIf(IsNull(rs("Nationality").value), "", Trim(rs("Nationality").value))
    TotOlds.Text = IIf(IsNull(rs("TotOlds").value), "", Trim(rs("TotOlds").value))
    Total.Text = IIf(IsNull(rs("Total").value), "", Trim(rs("Total").value))
    TotYoungs.Text = IIf(IsNull(rs("TotYoungs").value), "", Trim(rs("TotYoungs").value))
    Remark.Text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    TxtGroupName.Text = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))
    TxtReceptName.Text = IIf(IsNull(rs("ReceptName").value), "", Trim(rs("ReceptName").value))
    txtReceptID.Text = IIf(IsNull(rs("ReceptID").value), "", Trim(rs("ReceptID").value))
    TxtReceptOffice.Text = IIf(IsNull(rs("ReceptOffice").value), "", Trim(rs("ReceptOffice").value))
    TxtPArCode.Text = IIf(IsNull(rs("PArCode").value), "", Trim(rs("PArCode").value))
    TxtPassword.Text = IIf(IsNull(rs("Password").value), "", Trim(rs("Password").value))
    TxtSession.Text = IIf(IsNull(rs("Session").value), "", Trim(rs("Session").value))
    DcbPath.BoundText = IIf(IsNull(rs("PathID").value), 0, Trim(rs("PathID").value))
    TXtLargPrice.Text = IIf(IsNull(rs("LargPrice").value), 0, Trim(rs("LargPrice").value))
    TxtSmalPrice.Text = IIf(IsNull(rs("SmalPrice").value), 0, Trim(rs("SmalPrice").value))
    txtTotalPrice.Text = IIf(IsNull(rs("TotalPrice").value), 0, Trim(rs("TotalPrice").value))
    DcbVehicleType.BoundText = IIf(IsNull(rs("VehicleType").value), 0, Trim(rs("VehicleType").value))
    DcbLocation.BoundText = IIf(IsNull(rs("LocationID").value), 0, Trim(rs("LocationID").value))
    TxtCapacity.Text = IIf(IsNull(rs("Capacity").value), 0, Trim(rs("Capacity").value))
    TxtNoVehicle.Text = IIf(IsNull(rs("NoVehicle").value), 0, Trim(rs("NoVehicle").value))
    TxtPhone.Text = IIf(IsNull(rs("Phone").value), "", Trim(rs("Phone").value))
    CompanyID.BoundText = IIf(IsNull(rs("CompanyID").value), 0, Trim(rs("CompanyID").value))
    
    Dim ContactTime As Date
     If Not IsNull(rs("ReceptTime").value) Then
     ContactTime = FormatDateTime(rs("ReceptTime").value, vbShortTime)
      Me.ReceptTime.value = ContactTime
    End If
    FullGridData
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


Private Sub TxtSmalPrice_Change()
If Me.TxtModFlg.Text <> "R" Then
txtTotalPrice.Text = val(TXtLargPrice.Text) * val(TotOlds.Text) + val(TxtSmalPrice.Text) * val(TotYoungs.Text)
Total.Text = val(TotYoungs.Text) + val(TotOlds.Text)
End If
End Sub

Private Sub TxtSmalPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSmalPrice.Text, 0)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
   
    End If

    Select Case Index

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
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If Trim(BranchID.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Managerial Area"
            Else
                Msg = "Õœœ «·›—⁄ «Ê·« "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            BranchID.SetFocus
   '         SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text

           Case "N"
                rs.AddNew
                ID.Text = CStr(new_id("TblEndorseTrans", "ID", "", True))
           Case "E"
              Cn.Execute "Delete from  TblEndorseTransDet where EnTransID=" & val(ID.Text) & ""
                 
           End Select
        rs("ID").value = val(ID.Text)
        rs("SDate").value = SDate.value
        rs("RecordDateH").value = RecorddateH.value
        rs("BranchID").value = IIf(BranchID.BoundText = "", Null, BranchID.BoundText)
        rs("SeasonsID").value = IIf(SeasonsID.BoundText = "", Null, SeasonsID.BoundText)
        rs("ApproveID").value = IIf(ApproveID.Text = "", Null, ApproveID.Text)
        rs("DepandID").value = val(DcbDepandID.BoundText)
        'rs("FromCityID").value = IIf(FromCityID.BoundText = "", Null, (FromCityID.BoundText))
        'rs("GroupID").value = IIf(GroupID.BoundText = "", Null, GroupID.BoundText)
        rs("CompanyID").value = IIf(CompanyID.BoundText = "", Null, (CompanyID.BoundText))
        rs("Nationality").value = IIf(Nationality.BoundText = "", Null, (Nationality.BoundText))
        rs("TotOlds").value = IIf(TotOlds.Text = "", 0, val(TotOlds.Text))
        rs("TotYoungs").value = IIf(TotYoungs.Text = "", Null, val(TotYoungs.Text))
        rs("Total").value = IIf(Total.Text = "", Null, val(Total.Text))
        rs("Remark").value = IIf(Remark.Text = "", Null, Remark.Text)
        rs("creationdate").value = Date
        rs("creationuserID").value = user_id
        rs("TotalPrice").value = val(txtTotalPrice.Text)
        rs("SmalPrice").value = val(TxtSmalPrice.Text)
        rs("LargPrice").value = val(TXtLargPrice.Text)
        rs("PathID").value = val(DcbPath.BoundText)
        rs("Session").value = TxtSession.Text
        rs("Password").value = TxtPassword.Text
        rs("PArCode").value = TxtPArCode.Text
        rs("ReceptOffice").value = TxtReceptOffice.Text
        rs("ReceptID").value = txtReceptID.Text
        rs("ReceptName").value = TxtReceptName.Text
        rs("GroupName").value = TxtGroupName.Text
        rs("VehicleType").value = IIf(DcbVehicleType.BoundText = "", Null, DcbVehicleType.BoundText)
        rs("LocationID").value = IIf(DcbLocation.BoundText = "", Null, DcbLocation.BoundText)
        rs("Capacity").value = val(TxtCapacity.Text)
        rs("NoVehicle").value = val(TxtNoVehicle.Text)
        rs("Phone").value = TxtPhone.Text
        rs("ReceptTime").value = FormatDateTime(ReceptTime.value, vbShortTime)
        rs.update
    Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblEndorseTransDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("CarID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("EnTransID").value = val(ID.Text)
                RsDevsub("CarID").value = IIf((.TextMatrix(i, .ColIndex("CarID"))) = "", Null, val(.TextMatrix(i, .ColIndex("CarID"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("Capacity").value = IIf((.TextMatrix(i, .ColIndex("Capacity"))) = "", Null, val(.TextMatrix(i, .ColIndex("Capacity"))))
       RsDevsub.update
      End If
     Next i
    End With

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find " ID=" & val(ID.Text) & "", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Action()
  
        Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If ID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã  —ﬁ„ " & CHR(13)
        Msg = Msg + (ID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Confirm Delete ? " & CHR(13)
        Msg = Msg + (ID.Text) & CHR(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                                
                 
                   Cn.Execute "Delete from   TblEndorseTransDet where EnTransID=" & val(ID.Text) & ""
                StrSQL = "delete From TblEndorseTrans where  ID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From TblEndorseTrans "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   
                   
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                  
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
   
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
         Msg = "this process Not Aailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  «⁄ „«œ ‰ﬁ· ÕÃ«Ã ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «⁄ „«œ ‰ﬁ· ÕÃ«Ã" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «⁄ „«œ ‰ﬁ· ÕÃ«Ã" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«ÃÃ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ««⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «⁄ „«œ «—ﬂ«» «·ÕÃ«Ã", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub


