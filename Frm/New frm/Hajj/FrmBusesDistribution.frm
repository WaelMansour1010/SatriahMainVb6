VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBusesDistribution 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Ê“Ì⁄ Õ«›·«  «·„‘«⁄—"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   Icon            =   "FrmBusesDistribution.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   12105
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   12105
      _cx             =   21352
      _cy             =   9763
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
      Flags(0)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5115
         Left            =   -12660
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Width           =   12015
         _cx             =   21193
         _cy             =   9022
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
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1875
            Width           =   4275
         End
         Begin VB.TextBox TxtSession 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   2235
            Width           =   4275
         End
         Begin VB.TextBox TotOlds 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   8715
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   465
            Width           =   1575
         End
         Begin VB.TextBox TotYoungs 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   8715
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   825
            Width           =   1575
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   8715
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1170
            Width           =   1575
         End
         Begin VB.TextBox TXtLargPrice 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   465
            Width           =   1590
         End
         Begin VB.TextBox TxtSmalPrice 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   825
            Width           =   1590
         End
         Begin VB.TextBox TxtTotalPrice 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6555
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1170
            Width           =   1590
         End
         Begin VB.TextBox TxtGroupName 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1530
            Width           =   4275
         End
         Begin VB.TextBox TxtPassword 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   6555
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2235
            Width           =   3735
         End
         Begin VB.TextBox TxtPArCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1875
            Width           =   3735
         End
         Begin VB.TextBox Remark 
            Alignment       =   1  'Right Justify
            Height          =   795
            Left            =   270
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   2700
            Width           =   10020
         End
         Begin MSDataListLib.DataCombo SeasonsID 
            Height          =   315
            Left            =   270
            TabIndex        =   42
            Top             =   825
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPath 
            Height          =   315
            Left            =   270
            TabIndex        =   43
            Top             =   120
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID 
            Height          =   315
            Left            =   270
            TabIndex        =   44
            Top             =   1170
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Nationality 
            Height          =   315
            Left            =   6555
            TabIndex        =   45
            Top             =   1530
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   1200
            Left            =   135
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   3510
            Width           =   11715
            _cx             =   20664
            _cy             =   2117
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
               Format          =   94961666
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
            Left            =   270
            TabIndex        =   70
            Top             =   465
            Width           =   4275
            _ExtentX        =   7541
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
            Caption         =   "«·›∆…"
            Height          =   300
            Index           =   6
            Left            =   4995
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   465
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   165
            Index           =   10
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   825
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·œÊ—…"
            Height          =   285
            Index           =   12
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   2235
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«⁄ „«œ"
            Height          =   180
            Index           =   13
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1875
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œÿ «·”Ì—"
            Height          =   285
            Index           =   14
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂ»«—"
            Height          =   285
            Index           =   9
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«‰’«›"
            Height          =   150
            Index           =   5
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   825
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "««·„Ã„Ê⁄"
            Height          =   285
            Index           =   1
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1170
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄œœ"
            Height          =   390
            Index           =   15
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   390
            Index           =   16
            Left            =   6885
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ƒ””…"
            Height          =   180
            Index           =   7
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1170
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ã„Ê⁄…"
            Height          =   165
            Index           =   17
            Left            =   5025
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1530
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ã‰”Ì…"
            Height          =   270
            Index           =   0
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1530
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·”—Ì"
            Height          =   405
            Index           =   11
            Left            =   10830
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   2235
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·»«—ﬂÊœ"
            Height          =   285
            Index           =   18
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1875
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   405
            Index           =   3
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   2925
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5115
         Left            =   45
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   45
         Width           =   12015
         _cx             =   21193
         _cy             =   9022
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
            Height          =   270
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   4695
            Width           =   1995
         End
         Begin VB.TextBox TxtNoVehicle 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   4695
            Width           =   1980
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   4410
            Left            =   135
            TabIndex        =   73
            Top             =   120
            Width           =   11775
            _cx             =   20770
            _cy             =   7779
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
            FormatString    =   $"FrmBusesDistribution.frx":038A
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
            Height          =   255
            Left            =   9705
            TabIndex        =   84
            ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
            Top             =   4695
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
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
            ButtonImage     =   "FrmBusesDistribution.frx":04FA
            ButtonImageDisabled=   "FrmBusesDistribution.frx":6D5C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   255
            Left            =   7680
            TabIndex        =   85
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   4695
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   450
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
            ButtonImage     =   "FrmBusesDistribution.frx":25F46
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·”⁄…"
            Height          =   315
            Index           =   27
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   4695
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Õ«›·« "
            Height          =   315
            Index           =   26
            Left            =   5550
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   4695
            Width           =   1470
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic7 
      Height          =   750
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7680
      Width           =   12045
      _cx             =   21246
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
         Left            =   9015
         TabIndex        =   2
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":2C7A8
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
         Left            =   7995
         TabIndex        =   3
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":3300A
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
         Left            =   7005
         TabIndex        =   4
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":3986C
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
         Left            =   6015
         TabIndex        =   5
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":400CE
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
         Left            =   4965
         TabIndex        =   6
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":46930
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
         Left            =   2025
         TabIndex        =   7
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
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
         ButtonImage     =   "FrmBusesDistribution.frx":4D192
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
         Left            =   960
         TabIndex        =   8
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
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
         ButtonImage     =   "FrmBusesDistribution.frx":76DB4
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
         Left            =   3990
         TabIndex        =   9
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":7D616
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
         Left            =   2940
         TabIndex        =   10
         Top             =   120
         Width           =   960
         _ExtentX        =   1693
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
         ButtonImage     =   "FrmBusesDistribution.frx":83E78
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
         Left            =   1110
         TabIndex        =   86
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
         ButtonImage     =   "FrmBusesDistribution.frx":8A6DA
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
         Left            =   1545
         TabIndex        =   87
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
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
         ButtonImage     =   "FrmBusesDistribution.frx":90F3C
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
      Width           =   12045
      _cx             =   21246
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
      Caption         =   " Ê“Ì⁄ Õ«›·«  «·„‘«⁄—"
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
         ButtonImage     =   "FrmBusesDistribution.frx":9779E
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
         ButtonImage     =   "FrmBusesDistribution.frx":97B38
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
         ButtonImage     =   "FrmBusesDistribution.frx":97ED2
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
         ButtonImage     =   "FrmBusesDistribution.frx":9826C
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
      Width           =   12120
      _cx             =   21378
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
         Left            =   6870
         TabIndex        =   69
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " »Ê«”ÿ…"
         Height          =   330
         Index           =   29
         Left            =   10320
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   330
         Left            =   3945
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   120
         Width           =   930
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   330
         Index           =   2
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   330
         Index           =   4
         Left            =   975
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   1140
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   855
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   720
      Width           =   12075
      _cx             =   21299
      _cy             =   1508
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
      Begin VB.TextBox TxtOrderNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   165
         Width           =   1815
      End
      Begin VB.TextBox ID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9285
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   165
         Width           =   1470
      End
      Begin MSComCtl2.DTPicker SDate 
         Height          =   315
         Left            =   6630
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   94961667
         CurrentDate     =   37140
      End
      Begin MSDataListLib.DataCombo BranchID 
         Height          =   315
         Left            =   4950
         TabIndex        =   27
         Top             =   525
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Dynamic_Byte.NourHijriCal RecordDateH 
         Height          =   315
         Left            =   4980
         TabIndex        =   88
         Top             =   165
         Width           =   1515
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï «⁄ „«œ «—ﬂ«» «·„‘«⁄— —ﬁ„"
         Height          =   330
         Index           =   28
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   585
         Index           =   24
         Left            =   11325
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8190
         TabIndex        =   26
         Top             =   165
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”·”·"
         Height          =   465
         Index           =   8
         Left            =   10575
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   165
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmBusesDistribution"
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
  Dim Sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
Sql = "SELECT     dbo.TblBusesDistributionDet.ID, dbo.TblBusesDistributionDet.BusDistID, dbo.TblBusesDistributionDet.EmpID,dbo.TblBusesDistributionDet.Capacity, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
Sql = Sql & "                       dbo.TblEmployee.Emp_Namee, dbo.TblBusesDistributionDet.CarID, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
Sql = Sql & "                       dbo.TblCarsData.BoardNO , dbo.TblCarsData.Model, dbo.TblCarsData.OperatorN"
Sql = Sql & "  FROM         dbo.TBLCarTypes RIGHT OUTER JOIN"
Sql = Sql & "                       dbo.TblCarsData ON dbo.TBLCarTypes.id = dbo.TblCarsData.CarsTypeId RIGHT OUTER JOIN"
Sql = Sql & "                       dbo.TblBusesDistributionDet ON dbo.TblCarsData.id = dbo.TblBusesDistributionDet.CarID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblEmployee ON dbo.TblBusesDistributionDet.EmpID = dbo.TblEmployee.Emp_ID"
Sql = Sql & "  Where (dbo.TblBusesDistributionDet.BusDistID =" & val(ID.Text) & ") "


  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
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
            ID.Text = CStr(new_id("TblBusesDistribution", "ID", "", True))
         
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
GridInstallments.Rows = GridInstallments.Rows + 1
            TxtModFlg.Text = "E"
  
        Case 2

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Action

        
        Case 6
                Unload Me
         Case 7
                print_report2
         Case 9
         Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "Distribution"
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
ShowAttachments ID.Text, "20911201605"
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
   Dcombos.GetUsers DcbUser
End Sub


Private Sub Form_Load()
 '   On Error GoTo ErrTrap
        Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  „·› «·ÕÃ  "
    LogTextE = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""
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
            StrSQL = "SELECT  *  From TblBusesDistribution    "
    Else
            StrSQL = "SELECT  *  From TblBusesDistribution"
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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«   Ê“Ì⁄ Õ«›·«  «·„‘«⁄—    "
    LogTextE = " Exit Window " & "  "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

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
     Dim capacity As Integer
    IntCounter = 0
    capacity = 0
    With GridInstallments

        For i = .FixedRows To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("CarID"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            capacity = capacity + val(.TextMatrix(i, .ColIndex("Capacity")))
            End If
        Next i
 
    End With
    TxtNoVehicle.Text = IntCounter
      TxtCapacity.Text = capacity

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
                If SystemOptions.usertype <> UserAdminAll Then
               StrSQL = StrSQL & " and e.BranchID=" & Current_branch & ""
                End If
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


Private Sub ISButton3_Click()
On Error Resume Next
    With Me.GridInstallments
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
        ReLineGrid
    End With
End Sub

Private Sub ISButton4_Click()
 On Error Resume Next
 Me.GridInstallments.Clear flexClearScrollable, flexClearEverything
 Me.GridInstallments.Rows = 1
 ReLineGrid
End Sub

Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    SDate.value = ToGregorianDate(RecordDateH.value)
    End If
End Sub

Private Sub SDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(SDate.value)
End If
End Sub


Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   Ê“Ì⁄ Õ«›·«  «·„‘«⁄—"
            Else
                Me.Caption = "Hajj  Data"
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
                Me.Caption = "»Ì«‰«   Ê“Ì⁄ Õ«›·«  «·„‘«⁄— ( ÃœÌœ )"
            Else
                Me.Caption = "Buses Distribution(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄— ( ÃœÌœ )"
            Else
                Me.Caption = "Buses Distribution Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
           ' pnlHeader.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«      Ê“Ì⁄ Õ«›·«  «·„‘«⁄— (  ⁄œÌ· )"
            Else
                Me.Caption = "Buses Distribution Data(Edit)"
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
MySQL = "SELECT     TOP 100 PERCENT TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TblBusesDistribution.SDate AS DistSDate, "
MySQL = MySQL & "                       dbo.TblBusesDistribution.RecordDateH AS DistRecordDateH, dbo.TblBusesDistribution.NoVehicle AS DistNoVehicle,"
MySQL = MySQL & "                       dbo.TblBusesDistribution.Capacity AS DistCapacity, dbo.TblBusesDistribution.OrderNo, dbo.TblBusesDistribution.BranchID AS DIstBranchID,"
MySQL = MySQL & "                       TblBranchesData_1.branch_name AS Distbranch_name, TblBranchesData_1.branch_namee AS Distbranch_nameE,"
MySQL = MySQL & "                       dbo.TblBusesDistributionDet.Capacity AS DistCapacityDet, dbo.TblBusesDistributionDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblCarsData.CarsTypeId,"
MySQL = MySQL & "                       dbo.TblCarsData.Model, dbo.TblBusesDistributionDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TBLCarTypes.name AS CarType,"
MySQL = MySQL & "                       dbo.TBLCarTypes.namee AS CarTypeE, dbo.TblBusesDistribution.ID"
MySQL = MySQL & "  FROM         dbo.TBLCarTypes RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCarsData ON dbo.TBLCarTypes.id = dbo.TblCarsData.CarsTypeId RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBusesDistributionDet ON dbo.TblEmployee.Emp_ID = dbo.TblBusesDistributionDet.EmpID ON"
MySQL = MySQL & "                       dbo.TblCarsData.id = dbo.TblBusesDistributionDet.CarID RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBusesDistribution ON TblBranchesData_1.branch_id = dbo.TblBusesDistribution.BranchID ON"
MySQL = MySQL & "                       dbo.TblBusesDistributionDet.BusDistID = dbo.TblBusesDistribution.ID"
MySQL = MySQL & "  Where (dbo.TblBusesDistribution.orderNo <> 0) And (Not (dbo.TblBusesDistribution.orderNo Is Null)) And (dbo.TblBusesDistribution.ID=" & val(ID.Text) & " )"


     If SystemOptions.UserInterface = ArabicInterface Then
     
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetecting 2.rpt"
     Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetecting 2.rpt"
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
    RecordDateH.value = IIf(IsNull(rs("RecordDateH").value), ToHijriDate(Date), rs("RecordDateH").value)
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    BranchID.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    TxtOrderNo.Text = IIf(IsNull(rs("OrderNo").value), "", Trim(rs("OrderNo").value))
    TxtNoVehicle.Text = IIf(IsNull(rs("NoVehicle").value), "", Trim(rs("NoVehicle").value))
    TxtCapacity.Text = IIf(IsNull(rs("Capacity").value), "", Trim(rs("Capacity").value))
    FullGridData
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub



Private Sub TxtOrderNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtOrderNo.Text, 0)
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
                ID.Text = CStr(new_id("TblBusesDistribution", "ID", "", True))
           Case "E"
              Cn.Execute "Delete from  TblBusesDistributionDet where BusDistID=" & val(ID.Text) & ""
                 
           End Select
        rs("ID").value = val(ID.Text)
        rs("SDate").value = SDate.value
        rs("RecordDateH").value = RecordDateH.value
        rs("BranchID").value = IIf(BranchID.BoundText = "", Null, BranchID.BoundText)
        rs("OrderNo").value = val(TxtOrderNo.Text)
        rs("Capacity").value = val(TxtCapacity.Text)
        rs("NoVehicle").value = val(TxtNoVehicle.Text)
        rs.update
    Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBusesDistributionDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("CarID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("BusDistID").value = val(ID.Text)
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
                    Msg = "  „ Õ›Ÿ »Ì«‰«   Ê“Ì⁄ Õ«›·«  «·„‘«⁄— " & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & Chr(13)
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
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
        Msg = "”Ì „ Õ–› »Ì«‰«   Ê“Ì⁄ Õ«›·«  «·„‘«⁄—  —ﬁ„ " & Chr(13)
        Msg = Msg + (ID.Text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Confirm Delete ? " & Chr(13)
        Msg = Msg + (ID.Text) & Chr(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                                
                 
                   Cn.Execute "Delete from   TblBusesDistributionDet where BusDistID=" & val(ID.Text) & ""
                StrSQL = "delete From TblBusesDistribution where  ID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From TblBusesDistribution "
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ…  Ê“Ì⁄ Õ«›·«  «·„‘«⁄— "
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«   Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄— ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–«  Ê“Ì⁄ Õ«›·«  «·„‘«⁄—" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰   Ê“Ì⁄ Õ«›·«  «·„‘«⁄—" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«    Ê“Ì⁄ Õ«›·«  «·„‘«⁄—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub


