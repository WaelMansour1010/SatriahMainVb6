VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReportsNew 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… «· ﬁ«—Ì— "
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   Icon            =   "FrmReportsNew.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   10710
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   4125
      Left            =   30
      TabIndex        =   32
      Top             =   1620
      Width           =   6375
      _cx             =   11245
      _cy             =   7276
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   14871017
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "1|2|3|4|5|6|7|8|9"
      Align           =   0
      CurrTab         =   8
      FirstTab        =   0
      Style           =   3
      Position        =   0
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   8
         Left            =   45
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   7
         Left            =   -6930
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   6
         Left            =   -7230
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   5
         Left            =   -7530
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·⁄„Ì·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1035
            Left            =   570
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   570
            Width           =   5535
            Begin VB.TextBox TxtCusID 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   270
               Width           =   825
            End
            Begin MSDataListLib.DataCombo DcboCustomersSuppliers2 
               Height          =   315
               Left            =   750
               TabIndex        =   129
               Top             =   630
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   135
               Top             =   600
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   556
               ButtonStyle     =   1
               Caption         =   "..."
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
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·⁄„Ì·"
               Height          =   315
               Index           =   22
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   300
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄„Ì·"
               Height          =   315
               Index           =   12
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   660
               Width           =   885
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·’‰›"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   945
            Index           =   9
            Left            =   570
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   1620
            Width           =   5535
            Begin ImpulseButton.ISButton Cmd 
               Height          =   315
               Index           =   1
               Left            =   60
               TabIndex        =   133
               Top             =   540
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   556
               ButtonStyle     =   1
               Caption         =   "..."
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
               DrawFocusRectangle=   0   'False
            End
            Begin VB.TextBox TxtItemCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   210
               Width           =   2535
            End
            Begin MSDataListLib.DataCombo DcboItemName 
               Height          =   315
               Left            =   750
               TabIndex        =   125
               Top             =   540
               Width           =   3885
               _ExtentX        =   6853
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·’‰›"
               Height          =   315
               Index           =   14
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   540
               Width           =   795
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   315
               Index           =   13
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   210
               Width           =   915
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·—»Õ "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Index           =   8
            Left            =   570
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   2580
            Width           =   5535
            Begin VB.ComboBox Cbo 
               Height          =   315
               Index           =   0
               Left            =   2310
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   300
               Width           =   3075
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               Height          =   345
               Index           =   1
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   660
               Width           =   705
            End
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               Height          =   345
               Index           =   2
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   300
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               Height          =   225
               Index           =   21
               Left            =   570
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   720
               Width           =   285
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               Height          =   225
               Index           =   20
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   360
               Width           =   285
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   345
               Index           =   19
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   660
               Width           =   435
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ "
               Height          =   345
               Index           =   18
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   300
               Width           =   435
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   4
         Left            =   -7830
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
         Begin VB.Frame FraReportOptions 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  Œ«’… "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   3645
            Index           =   2
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   60
            Width           =   6135
            Begin VB.ComboBox CboTransactions 
               Height          =   315
               Left            =   1380
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   113
               Top             =   360
               Width           =   3555
            End
            Begin MSDataListLib.DataCombo DcboGroups1 
               Height          =   315
               Left            =   1380
               TabIndex        =   112
               Top             =   720
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄…"
               Height          =   315
               Index           =   10
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   750
               Width           =   1065
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
               Height          =   315
               Index           =   11
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   390
               Width           =   1065
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   3
         Left            =   -8130
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
         Begin VB.Frame FraReportOptions 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  Œ«’… „⁄ «· ﬁ—Ì—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   3135
            Index           =   0
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   -30
            Width           =   5505
            Begin VB.TextBox TxtToID 
               Alignment       =   2  'Center
               Height          =   345
               Index           =   0
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   570
               Width           =   1065
            End
            Begin VB.TextBox TxtFromID 
               Alignment       =   2  'Center
               Height          =   345
               Index           =   0
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   600
               Width           =   1065
            End
            Begin VB.TextBox TxtTransSerial 
               Alignment       =   1  'Right Justify
               Height          =   345
               Index           =   0
               Left            =   1230
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   2130
               Width           =   1425
            End
            Begin VB.ComboBox CboTrans 
               Height          =   315
               Index           =   1
               Left            =   1230
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   1770
               Width           =   3045
            End
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               Height          =   345
               Index           =   0
               Left            =   3060
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1380
               Width           =   1245
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   19
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   1350
               Width           =   1545
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   0
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  ToolTipText     =   "«ﬂ»— „‰"
                  Top             =   0
                  Width           =   465
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "="
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   1
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  ToolTipText     =   "Ì”«ÊÏ"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   2
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  ToolTipText     =   "«’€— „‰"
                  Top             =   0
                  Width           =   555
               End
            End
            Begin MSDataListLib.DataCombo DCboPaymentClient 
               Height          =   315
               Left            =   1230
               TabIndex        =   96
               Top             =   240
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboUsers 
               Height          =   315
               Index           =   0
               Left            =   1230
               TabIndex        =   97
               Top             =   2550
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Index           =   1
               Left            =   1230
               TabIndex        =   98
               Top             =   1020
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄„Ì· √Ê «·„Ê—œ"
               Height          =   405
               Index           =   2
               Left            =   4290
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   210
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï —ﬁ„"
               Height          =   375
               Index           =   54
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   630
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ —ﬁ„"
               Height          =   375
               Index           =   53
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   690
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„” Œœ„"
               Height          =   375
               Index           =   52
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   2580
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂ… «Ê «·›« Ê—…"
               Height          =   375
               Index           =   45
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   2190
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰« Ã ⁄‰"
               Height          =   285
               Index           =   44
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1830
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·„»·€"
               Height          =   375
               Index           =   43
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   1410
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   375
               Index           =   42
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   1050
               Width           =   915
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   2
         Left            =   -8430
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
         Begin VB.Frame FraReportOptions 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  Œ«’… „⁄  ﬁ«—Ì— «·„’—Ê›« "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   3135
            Index           =   1
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   90
            Width           =   6135
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   375
               Index           =   21
               Left            =   1860
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   1590
               Width           =   1545
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   8
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  ToolTipText     =   "«’€— „‰"
                  Top             =   0
                  Width           =   555
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "="
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   7
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  ToolTipText     =   "Ì”«ÊÏ"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Index           =   6
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  ToolTipText     =   "«ﬂ»— „‰"
                  Top             =   0
                  Width           =   465
               End
            End
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               Height          =   345
               Index           =   1
               Left            =   3690
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   1590
               Width           =   1245
            End
            Begin VB.TextBox TxtFromID 
               Alignment       =   2  'Center
               Height          =   345
               Index           =   2
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   780
               Width           =   1065
            End
            Begin VB.TextBox TxtToID 
               Alignment       =   2  'Center
               Height          =   345
               Index           =   2
               Left            =   1890
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   780
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo DcboUsers1 
               Height          =   315
               Left            =   1890
               TabIndex        =   75
               Top             =   1980
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox1 
               Height          =   315
               Index           =   1
               Left            =   1890
               TabIndex        =   76
               Top             =   1230
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboExpensesName 
               Height          =   315
               Left            =   1890
               TabIndex        =   77
               Top             =   360
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   375
               Index           =   67
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1260
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·„»·€"
               Height          =   375
               Index           =   66
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1620
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„” Œœ„"
               Height          =   375
               Index           =   65
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   2010
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ —ﬁ„"
               Height          =   375
               Index           =   64
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   840
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï —ﬁ„"
               Height          =   285
               Index           =   63
               Left            =   3060
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   810
               Width           =   675
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„’—Ê›« "
               Height          =   315
               Index           =   17
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   420
               Width           =   1065
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   1
         Left            =   -8730
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
         Begin VB.Frame FraReportOptions 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  „⁄  ﬁ«—Ì— «·√’‰«›"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   3135
            Index           =   4
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   90
            Width           =   6165
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·—’Ìœ"
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
               Height          =   1185
               Index           =   5
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   660
               Width           =   4455
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√’‰«› «· Ï —’ÌœÂ« ’›— ›ﬁÿ"
                  Height          =   255
                  Index           =   3
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   240
                  Width           =   3855
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√’‰«› «· Ï —’ÌœÂ« «ﬁ· „‰ «·’›— ( —’Ìœ »«·”«·»)"
                  Height          =   315
                  Index           =   4
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   480
                  Width           =   3855
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·√—’œ… «·’›—Ì… Ê«·”«·»…"
                  Height          =   315
                  Index           =   5
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   750
                  Value           =   -1  'True
                  Width           =   3855
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—’Ìœ ’›—Ï »‰«¡ ⁄·Ï ...."
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
               Height          =   1185
               Index           =   6
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1860
               Width           =   4455
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·Õ«·« "
                  Height          =   315
                  Index           =   9
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   840
                  Value           =   -1  'True
                  Width           =   4155
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—’Ìœ ’›— ...·«‰ «·—’Ìœ «‰ ÂÏ „‰ «·„Œ“‰"
                  Height          =   315
                  Index           =   10
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   540
                  Width           =   4155
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—’Ìœ ’›— ...·«‰ «·√’‰«› ·„  œŒ· ›Ï «Ï ⁄„·Ì«   Ã«—Ì…"
                  Height          =   315
                  Index           =   11
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   240
                  Width           =   4155
               End
            End
            Begin MSDataListLib.DataCombo DcboGroups 
               Height          =   315
               Left            =   1590
               TabIndex        =   65
               Top             =   300
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄…"
               Height          =   315
               Index           =   1000
               Left            =   4650
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   300
               Width           =   1065
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3750
         Index           =   0
         Left            =   -9030
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   330
         Width           =   6285
         _cx             =   11086
         _cy             =   6615
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
         Begin VB.ComboBox CboSalesType 
            Height          =   315
            Left            =   2130
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1920
            Width           =   3045
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2250
            Width           =   1695
         End
         Begin VB.Frame FraDueOptions 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„·Ì«  «·√Ã·… Ê«· Õ’Ì·«  «· Ï «Ã—Ì  ⁄·ÌÂ«"
            ForeColor       =   &H00800000&
            Height          =   1455
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   2280
            Visible         =   0   'False
            Width           =   3465
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ »Ì«‰«  ⁄„·Ì«  «· Õ’Ì·"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   1140
               Width           =   2535
            End
            Begin VB.OptionButton OptDebitTrans 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— √Ã·…   „  Õ’Ì·Â« »«·ﬂ«„·"
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
               Height          =   255
               Index           =   4
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   240
               Width           =   3345
            End
            Begin VB.OptionButton OptDebitTrans 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— √Ã·… Ê „  Õ’Ì· Ã“¡ „‰Â«"
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
               Height          =   255
               Index           =   5
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   450
               Width           =   3345
            End
            Begin VB.OptionButton OptDebitTrans 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— √Ã·… Ê·„ ÌÕ’· „‰Â« «Ï ‘Ï¡"
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
               Height          =   255
               Index           =   6
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   660
               Width           =   3345
            End
            Begin VB.OptionButton OptDebitTrans 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ· «·›Ê« Ì— «·√Ã·…"
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
               Height          =   255
               Index           =   7
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   870
               Value           =   -1  'True
               Width           =   3345
            End
         End
         Begin VB.ComboBox CboPaymentMethod 
            Height          =   315
            Left            =   2130
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   240
            Width           =   3045
         End
         Begin MSDataListLib.DataCombo DcboBoxTrans 
            Height          =   315
            Left            =   2130
            TabIndex        =   45
            Top             =   600
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboStores 
            Height          =   315
            Left            =   2130
            TabIndex        =   46
            Top             =   930
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpID 
            Height          =   315
            Left            =   2130
            TabIndex        =   47
            Top             =   1260
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboCustomersSuppliers1 
            Height          =   315
            Left            =   2130
            TabIndex        =   48
            Top             =   1590
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   285
            Index           =   2
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   285
            Index           =   5
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   930
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   285
            Index           =   6
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   315
            Index           =   7
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   1590
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·»Ì⁄"
            Height          =   315
            Index           =   8
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1920
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï ﬁÌ„… «·›« Ê—…"
            Height          =   435
            Index           =   9
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   2250
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   17
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   270
            Width           =   915
         End
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ’„Ì„ «· ﬁ—Ì—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Index           =   4
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6810
      Width           =   5445
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬂ«‰ ⁄—÷ «·—”„ «·»Ì«‰Ï ›Ï «· ﬁ—Ì—"
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
         Height          =   555
         Index           =   7
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   750
         Width           =   3045
         Begin VB.OptionButton OptChart 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Â«Ì… «· ﬁ—Ì—"
            Height          =   315
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   210
            Width           =   1185
         End
         Begin VB.OptionButton OptChart 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»œ«Ì… «· ﬁ—Ì—"
            Height          =   315
            Index           =   0
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   210
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VB.CheckBox ChkChart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ —”„ »Ì«‰Ï ›Ï «· ﬁ—Ì—"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   3090
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   780
         Width           =   2235
      End
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ ŒÌ«—«   ÕœÌœ «·»Ì«‰«  ›Ï «· ﬁ—Ì—"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1320
         Width           =   3525
      End
      Begin VB.OptionButton OptSort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ’«⁄œÌ"
         Height          =   255
         Index           =   3
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   510
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton OptSort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ‰«“·Ì"
         Height          =   285
         Index           =   2
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   510
         Width           =   1275
      End
      Begin VB.ComboBox CboReportStyle 
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   180
         Width           =   4095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ Ã«Â «· — Ì»"
         Height          =   255
         Index           =   16
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   510
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ÿ«„ «·⁄—÷"
         Height          =   285
         Index           =   15
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   210
         Width           =   795
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " — Ì» «·»Ì«‰«  ›Ï «· ﬁ—Ì—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1005
      Index           =   3
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   5760
      Width           =   3855
      Begin VB.OptionButton OptSort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ‰«“·Ì"
         Height          =   255
         Index           =   1
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   660
         Width           =   1125
      End
      Begin VB.OptionButton OptSort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ’«⁄œÌ"
         Height          =   255
         Index           =   0
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   660
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.ComboBox CboSortData 
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   270
         Width           =   2505
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ Ã«Â «· — Ì»"
         Height          =   285
         Index           =   1
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " — Ì» »‰«¡ ⁄·Ï "
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ Ì«— «·› —… «·“„‰Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1005
      Index           =   0
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5760
      Width           =   2445
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   -36522
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   69
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   615
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   71
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê’› «· ﬁ—Ì—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   885
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   690
      Width           =   6375
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   645
         Index           =   4
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   6285
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «· ﬁ—Ì—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   585
      Index           =   2
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   6375
      Begin VB.TextBox TxtNodeReport 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   210
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   3
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   210
         Width           =   6285
      End
   End
   Begin MSComctlLib.TreeView TrvReports 
      Height          =   8835
      Left            =   6450
      TabIndex        =   0
      Top             =   30
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   15584
      _Version        =   393217
      Indentation     =   18
      Style           =   7
      Appearance      =   1
   End
   Begin ImpulseButton.ISButton CmdExit 
      Height          =   345
      Left            =   90
      TabIndex        =   10
      Top             =   8520
      Width           =   1245
      _ExtentX        =   2196
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
      ButtonImage     =   "FrmReportsNew.frx":038A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   345
      Index           =   0
      Left            =   2670
      TabIndex        =   11
      Top             =   8520
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "FrmReportsNew.frx":0724
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   345
      Index           =   7
      Left            =   1380
      TabIndex        =   12
      Top             =   8520
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "FrmReportsNew.frx":0ABE
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   345
      Index           =   1
      Left            =   3960
      TabIndex        =   13
      Top             =   8520
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "FrmReportsNew.frx":0E58
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6450
      X2              =   0
      Y1              =   8460
      Y2              =   8460
   End
End
Attribute VB_Name = "FrmReportsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo(5) As clsDCboSearch

Private Sub Cbo_Change(Index As Integer)

    Select Case Index

        Case 0

            If Me.Cbo(Index).ListIndex = -1 Then
                lbl(18).Enabled = False
                lbl(19).Enabled = False
                lbl(20).Enabled = False
                lbl(21).Enabled = False
                Me.Txt(1).Enabled = False
                Me.Txt(2).Enabled = False
            Else
                lbl(18).Enabled = True
                lbl(19).Enabled = True
                lbl(20).Enabled = True
                lbl(21).Enabled = True
                Me.Txt(1).Enabled = True
                Me.Txt(2).Enabled = True

                If Me.Cbo(0).ListIndex = 0 Then
                    lbl(20).Caption = "%"
                    lbl(21).Caption = "%"
                Else
                    lbl(20).Caption = ""
                    lbl(21).Caption = ""
                End If
            End If

    End Select

End Sub

Private Sub Cbo_Click(Index As Integer)
    Cbo_Change Index
End Sub

Private Sub CboPaymentMethod_Change()

    If Me.CboPaymentMethod.ListIndex = 1 Then
        Me.FraDueOptions.Visible = True
    Else
        Me.FraDueOptions.Visible = False
    End If

End Sub

Private Sub CboPaymentMethod_Click()
    CboPaymentMethod_Change
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DcboCustomersSuppliers2
            FrmCustemerSearch.Show vbModal

        Case 1
            Load FrmItemSearch
            FrmItemSearch.RetrunType = 1
            Set FrmItemSearch.DcboItems = Me.DcboItemName
            FrmItemSearch.Show vbModal
    End Select

End Sub

Private Sub CmdPrint_Click(Index As Integer)
    Dim m_PrintTarget As PrintTarget

    '------------------------------------
    If Index = 0 Then
        m_PrintTarget = PrinterTarget
    Else
        m_PrintTarget = WindowTarget
    End If

    '------------------------------------
    Select Case Me.TxtNodeReport.text

        Case "Report1"
            ' ﬁ—Ì— «·„»Ì⁄« 
            ShowSalesReport m_PrintTarget

        Case "Report27"
            ' ﬁ—Ì— «·„’—Ê›« 
            ShowExpensesReport m_PrintTarget

        Case "Report44"
            ' ﬁ«—Ì— «·√—’œ… «·’›—Ì… ··√’‰«›
            ShowZeroItemsStock m_PrintTarget

        Case "Report45"
            ' ﬁ—Ì— «·⁄„·Ì«  «· Ã«—Ì… ⁄·Ï „Ã„Ê⁄… „⁄Ì‰…
            ShowGroupsItemsTransactions m_PrintTarget

        Case "Report46"
            '⁄—÷  ﬁ«—Ì— «·√—»«Õ
            ShowProfits m_PrintTarget

        Case Else
    End Select

End Sub

Private Sub WriteCaption(LblCtrl As Label, _
                         StrCaption As String)
    Dim i As Long
    Dim LngStrLenth As Long
    Dim StrTemp As String
    LngStrLenth = Len(StrCaption)

    For i = 1 To LngStrLenth

        DoEvents
        StrTemp = Mid$(StrCaption, 1, i)
        LblCtrl.Caption = StrTemp

        DoEvents
        Sleep 5
    Next i

End Sub

Private Sub Form_Load()
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim ShowTax As Boolean
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Dim Msg As String
    Dim i As Integer

    On Error GoTo ErrTrap

    LoadTree
    '-------------------
    Set Dcombos = New ClsDataCombos
    'OpInx = 0
    Screen.MousePointer = vbArrowHourglass
    '---------ŒÌ«—«  «·Õ—ﬂ«  «· Ã«—Ì… «·‰ﬁœÌ…
    Dcombos.GetBoxes Me.DcboBoxTrans
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboBoxTrans
    Dcombos.GetStores Me.DcboStores
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboStores
    Dcombos.GetEmployees Me.DcboEmpID
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DcboEmpID

    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomersSuppliers1
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboCustomersSuppliers1

    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomersSuppliers2, True
    Set cSearchDcbo(4) = New clsDCboSearch
    Set cSearchDcbo(4).Client = Me.DcboCustomersSuppliers2
    cSearchDcbo(4).SetBuddyText TxtCusID

    Dcombos.GetItemsNames Me.DcboItemName
    Set cSearchDcbo(5) = New clsDCboSearch
    Set cSearchDcbo(5).Client = Me.DcboItemName

    With CboPaymentMethod
        .Clear
        .AddItem "‰ﬁœÏ"
        .AddItem "√Ã·"
        .AddItem "«·ﬂ·"
    End With

    With Me.CboSalesType
        .Clear
        .AddItem "ﬁÿ«⁄Ï"
        .AddItem " Ã«—Ï"
        .AddItem "«·ﬂ·"
    End With

    With Me.CboTransactions
        .Clear
        .AddItem "„»Ì⁄« "
        .AddItem "„‘ —Ì« "
        .AddItem "„— Ã⁄ «·„»Ì⁄« "
        .AddItem "„— Ã⁄ «·„‘ —Ì« "
    End With

    With Me.Cbo(0)
        .Clear
        .AddItem "‰”»… «·—»Õ"
        .AddItem "ﬁÌ„… «·—»Õ"
        .ListIndex = -1
    End With

    For i = Me.Ele.LBound To Me.Ele.UBound
        Me.Ele(i).BackColor = &HE2E9E9
    Next i

    Me.TabMain.TabHeight = 1
    'Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName
    'Dcombos.GetCustomersSuppliers 0, Me.DBCboClientNameBuy
    'Dcombos.GetCustomersSuppliers 0, Me.Dcbomaintenece
    'Dcombos.GetCustomersSuppliers 0, Me.DcboClientTax
    'Dcombos.GetCustomersSuppliers 0, Me.DCasingCust
    'Dcombos.GetCustomersSuppliers 0, Me.DCboPaymentClient
    'Dcombos.GetCustomersSuppliers 0, DCboInstallment, False
    'Dcombos.GetCustomersSuppliers 0, DcboCusRePurchase, True
    'Dcombos.GetItemSGroups Me.DCboGroup, False
    Dcombos.GetItemSGroups Me.DcboGroups1, False
    '--------------------------------------
    'Dcombos.GetStores DCboStoreName
    'Dcombos.GetStores DcboStores
    '---------------------------------
    '«·„’—Ê›« 
    'Dcombos.GetBoxes Me.DcboBox(1)
    'Dcombos.GetBoxes Me.DcboBox(2)
    'Dcombos.GetBoxes Me.DcboBox(3)
    '---------------------------------
    Dcombos.GetUsers Me.DcboUsers1 '«·„’—Ê›« 
    'Dcombos.GetUsers Me.DcboUsers(1)
    'Dcombos.GetUsers Me.DcboUsers(2)
    '---------------------------------
    'Dcombos.GetItemsNames DCboItem
    'Dcombos.GetItemsNames DCboItemBuy
    'Dcombos.GetItemsNames DCboItemName
    'Dcombos.GetItemsNames Item
    'Dcombos.GetItemsNames Me.DcboItemName1
    Dcombos.GetExpensesType Me.DCboExpensesName
    'Dcombos.GetCustomersSuppliers 1, DcboClient, False
    'Dcombos.GetCustomersSuppliers 2, DcboCompany, False
    'Dcombos.GetEmployees Me.DCmboEmp

    ''≈Œ›«¡ «·Ã“¡ «·Œ«’ »÷—«∆» «·„»Ì⁄« 
    'ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    'C1TabMain.TabVisible(5) = ShowTax
    'Set RsTemp = New ADODB.Recordset
    '
    '
    ''«·„ÊŸ›Ì‰
    ''ShowTip
    'Set cSearchDcbo(0) = New clsDCboSearch
    'Set cSearchDcbo(0).Client = Me.DBCboClientName
    '
    'Set cSearchDcbo(1) = New clsDCboSearch
    'Set cSearchDcbo(1).Client = Me.DBCboClientNameBuy
    '
    'Set cSearchDcbo(2) = New clsDCboSearch
    'Set cSearchDcbo(2).Client = Me.DcboClientTax
    '
    'Set cSearchDcbo(3) = New clsDCboSearch
    'Set cSearchDcbo(3).Client = Me.DCboItem
    '
    'Set cSearchDcbo(4) = New clsDCboSearch
    'Set cSearchDcbo(4).Client = Me.DCboItemBuy
    '
    'Set cSearchDcbo(5) = New clsDCboSearch
    'Set cSearchDcbo(5).Client = Me.Dcbomaintenece
    '
    'Set cSearchDcbo(6) = New clsDCboSearch
    'Set cSearchDcbo(6).Client = Me.DCboGroup
    '
    'Set cSearchDcbo(7) = New clsDCboSearch
    'Set cSearchDcbo(7).Client = Me.DCboStoreName
    '
    'Set cSearchDcbo(8) = New clsDCboSearch
    'Set cSearchDcbo(8).Client = Me.DCboItemName
    '
    'Set cSearchDcbo(9) = New clsDCboSearch
    'Set cSearchDcbo(9).Client = Me.DCboExpensesName
    '
    'Set cSearchDcbo(10) = New clsDCboSearch
    'Set cSearchDcbo(10).Client = Me.DCasingCust
    '
    'Set cSearchDcbo(11) = New clsDCboSearch
    'Set cSearchDcbo(11).Client = Me.DCboPaymentClient
    '
    'Set cSearchDcbo(12) = New clsDCboSearch
    'Set cSearchDcbo(12).Client = DCmboEmp
    '
    'Set cSearchDcbo(13) = New clsDCboSearch
    'Set cSearchDcbo(13).Client = DcboItemName1
    '
    'Set cSearchDcbo(14) = New clsDCboSearch
    'Set cSearchDcbo(14).Client = DcboClient
    '
    'Set cSearchDcbo(15) = New clsDCboSearch
    'Set cSearchDcbo(15).Client = DcboCompany
    '
    'Set cSearchDcbo(16) = New clsDCboSearch
    'Set cSearchDcbo(16).Client = DCboInstallment
    '
    'Resize_Form Me
    'With Me.CboTrans(0)
    '    .Clear
    '    .AddItem "«·„»Ì⁄« "
    '    .AddItem "«·„‘ —Ì« "
    'End With
    '
    'With Me.CboTrans(1)
    '    .Clear
    '    .AddItem "›Ê« Ì— «·‘—«¡"
    '    .AddItem "„— Ã⁄ «·„»Ì⁄« "
    '    .AddItem "«·ﬂ·"
    'End With
    'With Me.CboTrans(2)
    '    .Clear
    '    .AddItem "›« Ê—… »Ì⁄"
    '    .AddItem "›« Ê—… „‘ —Ì« "
    '    .AddItem "„— Ã⁄ „»Ì⁄« "
    '    .AddItem "„— Ã⁄ „‘ —Ì« "
    '    .AddItem "—’Ìœ ≈›  «ÕÏ"
    '    .AddItem "Ã—œ „Œ“‰"
    'End With
    'With Me.CboTrans(3)
    '    .Clear
    '    .AddItem "›Ê« Ì— «·»Ì⁄"
    '    .AddItem "„— Ã⁄ «·„‘ —Ì« "
    '    .AddItem "«·ﬂ·"
    '    '.AddItem "«·’Ì«‰…"
    '    '.AddItem "Œœ„« "
    'End With
    '
    'C1TabMain.CurrTab = 0
    'XPChk(2).Value = vbUnchecked
    'XPChk_Click 2
    '
    'OptMisc_Click (0)
    'OptPayType(2).Value = True
    'OptPayType_Click 2
    '
    'Chk(6).Value = vbUnchecked
    'Chk_Click 6
    'ClientOpt_Click (0)
    'SuplierOpt_Click (0)
    '
    'Me.HelpContextID = 150
    'Msg = "Ì„ﬂ‰ﬂ ≈Œ Ì«— «Ï „‰ «Ê ﬂ· Â–Â «·ŒÌ«—«  · ÕœÌœ ‘—Êÿ ⁄—÷  ﬁ—Ì— «·„œ›Ê⁄«  «·–Ï  —Ìœ ⁄—÷Â"
    'Me.lbl(51).Caption = Msg
    'Msg = "Ì„ﬂ‰ﬂ ≈Œ Ì«— «Ï „‰ «Ê ﬂ· Â–Â «·ŒÌ«—«  · ÕœÌœ ‘—Êÿ ⁄—÷  ﬁ—Ì— «·„ﬁ»Ê÷«  «·–Ï  —Ìœ ⁄—÷Â"
    'Me.lbl(55).Caption = Msg
    '
    'Msg = " ÕœÌœ «·› —… «· «—ÌŒÌ… ÌﬂÊ‰ »Œ’Ê’  «—ÌŒ ≈” Õﬁ«ﬁ «·ﬁÌ„… «·√Ã·… ⁄·Ï «·‘—ﬂ…"
    'lbl(45).Caption = Msg
    'Msg = " ÕœÌœ «·› —… «· «—ÌŒÌ… ÌﬂÊ‰ »Œ’Ê’  «—ÌŒ ≈” Õﬁ«ﬁ «·ﬁÌ„… «·√Ã·… ··‘—ﬂ…"
    'lbl(46).Caption = Msg
    'Msg = " ÕœÌœ «·› —… «· «—ÌŒÌ… ÌﬂÊ‰ »Œ’Ê’  «—ÌŒ  Õ’Ì· «·‘Ìﬂ"
    'lbl(47).Caption = Msg
    'Msg = " ÕœÌœ «·› —… «· «—ÌŒÌ… ÌﬂÊ‰ »Œ’Ê’  «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ"
    'lbl(48).Caption = Msg
    'Msg = " ÕœÌœ «·› —… «· «—ÌŒÌ… ÌﬂÊ‰ »Œ’Ê’  «—ÌŒ  ”œÌœ «·‘Ìﬂ"
    'lbl(49).Caption = Msg
    'Msg = " ÕœÌœ «·› —… «· «—ÌŒÌ… ÌﬂÊ‰ »Œ’Ê’  «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ"
    'lbl(50).Caption = Msg
    'With Me.CboCusBalanceType(0)
    '    .Clear
    '    .AddItem "‰Ÿ«„ «·œ«∆‰ Ê«·„œÌ‰"
    '    .AddItem "‰Ÿ«„   «·Ï «·⁄„·Ì«  »«· ”·”·"
    '    .AddItem "‰Ÿ«„ «·œ«∆‰ Ê«·„œÌ‰(»«·≈÷«›… ≈·Ï ⁄—÷ «·√’‰«›)"
    '    .ListIndex = 0
    'End With
    'With Me.CboCusBalanceType(1)
    '    .Clear
    '    .AddItem "‰Ÿ«„ «·œ«∆‰ Ê«·„œÌ‰"
    '    .AddItem "‰Ÿ«„   «·Ï «·⁄„·Ì«  »«· ”·”·"
    '    .AddItem "‰Ÿ«„ «·œ«∆‰ Ê«·„œÌ‰(»«·≈÷«›… ≈·Ï ⁄—÷ «·√’‰«›)"
    '    .ListIndex = 0
    'End With
    Screen.MousePointer = vbDefault
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
    '-------------------
    HideFraOptions
    '-------------------
    Resize_Form Me
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadTree()
    Dim XNode As MSComctlLib.Node
    Dim i As Long

    'Report46
    Make_RightToLeft Me.TrvReports

    With Me.TrvReports
        .Appearance = ccFlat
        .BorderStyle = ccFixedSingle
        .LabelEdit = tvwManual
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        Set .ImageList = mdifrmmain.ImgLstTree
        Set XNode = .Nodes.Add(, , "Root", " ﬁ«—Ì— »—‰«„Ã œÌ‰«„Ìﬂ »«Ì  «·„ ﬂ«„·Ï", "DReport", "DReport")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group1", " ﬁ«—Ì— «·„»Ì⁄« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group1", tvwChild, "Report1", " ﬁ—Ì— »«·„»Ì⁄« ", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group2", " ﬁ«—Ì— «·„‘ —Ì« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group2", tvwChild, "Report4", " ﬁ—Ì— »«·„‘ —Ì«  «·‰ﬁœÌ…", "Purchase", "Purchase")
            
        Set XNode = .Nodes.Add("Root", tvwChild, "Group3", " ﬁ«—Ì— „— Ã⁄ «·„»Ì⁄« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group3", tvwChild, "Report7", " ﬁ—Ì— »„— Ã⁄ «·„»Ì⁄«  «·‰ﬁœÌ…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group4", " ﬁ«—Ì— „— Ã⁄ «·„‘ —Ì« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group4", tvwChild, "Report10", " ﬁ—Ì— »„— Ã⁄ «·„‘ —Ì«  «·‰ﬁœÌ…", "Purchase", "Purchase")
            
        Set XNode = .Nodes.Add("Root", tvwChild, "Group5", " ﬁ«—Ì— «·√’‰«›", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group5", tvwChild, "Report13", " ﬁ—Ì— »Ì«‰«  «·√’‰«›", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group5", tvwChild, "Report14", " ﬁ—Ì— »«·√’‰«› «· Ï »·€  Õœ «·ÿ·»", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group5", tvwChild, "Report15", " ﬁ—Ì— »√Ï ⁄„·Ì… ⁄·Ï «·’‰›", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group5", tvwChild, "Report16", " ﬁ—Ì— ﬂ«—  «·’‰›", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group5", tvwChild, "Report44", " ﬁ—Ì— »«·√—’œ… «·’›—Ì… Ê«·”«·»… „‰ «·√’‰«›", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group5", tvwChild, "Report45", " ﬁ—Ì— »«·⁄„·Ì«  «· Ã«—Ì… ⁄·Ï „Ã„Ê⁄… √’‰«›", "Purchase", "Purchase")
            
        Set XNode = .Nodes.Add("Root", tvwChild, "Group6", " ﬁ«—Ì— «·⁄„·«¡", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group6", tvwChild, "Report17", "ÿ»«⁄… »Ì«‰«  «·⁄„·«¡", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group6", tvwChild, "Report18", " ﬁ—Ì— ﬂ‘› Õ”«» ⁄„Ì·", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group6", tvwChild, "Report19", " ﬁ—Ì— »«·⁄„·Ì«  «· Ã«—Ì… „⁄ «·⁄„Ì·", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group6", tvwChild, "Report20", " ﬁ—Ì— »„‘ —Ì«  ⁄„Ì· „‰ «·√’‰«›", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group6", tvwChild, "Report21", " ﬁ—Ì— »„‘ —Ì«  ⁄„Ì· „‰ ’‰› „Õœœ", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group7", " ﬁ«—Ì— «·„Ê—œÌ‰", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group7", tvwChild, "Report22", "ÿ»«⁄… »Ì«‰«  «·„Ê—œÌ‰", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group7", tvwChild, "Report23", " ﬁ—Ì— ﬂ‘› Õ”«» „Ê—œ", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group7", tvwChild, "Report24", " ﬁ—Ì— »«·⁄„·Ì«  «· Ã«—Ì… „⁄ «·„Ê—œ", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group7", tvwChild, "Report25", " ﬁ—Ì— »«·„‘ —Ì«  „‰ «·„Ê—œ „‰ «·√’‰«›", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group7", tvwChild, "Report26", " ﬁ—Ì— »«·„‘ —Ì«  „‰  „Ê—œ „‰ ’‰› „Õœœ", "Purchase", "Purchase")
            
        Set XNode = .Nodes.Add("Root", tvwChild, "Group8", " ﬁ«—Ì— «·„’—Ê›« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group8", tvwChild, "Report27", " ﬁ«—Ì— «·„’—Ê›« ", "Purchase", "Purchase")
            
        Set XNode = .Nodes.Add("Root", tvwChild, "Group9", " ﬁ«—Ì— «·„œ›Ê⁄« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group9", tvwChild, "Report28", " ﬁ«—Ì— «·„œ›Ê⁄« ", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group10", " ﬁ«—Ì— «·„ﬁ»Ê÷« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group10", tvwChild, "Report29", " ﬁ«—Ì— «·„ﬁ»Ê÷« ", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group11", " ﬁ«—Ì— «·„⁄«„·«  «·„«·Ì…", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report30", " ﬁ—Ì— »«·„»«·€ «·„«·Ì… «·√Ã·… ⁄·Ï «·‘—ﬂ…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report31", " ﬁ—Ì— »«·„»«·€ «·„«·Ì… «·√Ã·… ⁄·Ï «·⁄„·«¡ Ê«·„Ê—œÌ‰", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report32", " ﬁ—Ì— »«·‘Ìﬂ«   „  Õ’Ì·Â«", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report33", " ﬁ—Ì— »«·‘Ìﬂ«   Õ  «· Õ’Ì·", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report34", " ﬁ—Ì— »«·‘Ìﬂ«   „  ”œÌœÂ«", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report35", " ﬁ—Ì— »«·‘Ìﬂ«  „ÿ·Ê»  ”œÌœÂ«", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report36", "  ﬁ—Ì— »«·Œ’Ê„«   «·„”„ÊÕ…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report37", "  ﬁ—Ì— »«·Œ’Ê„«  «·„ﬂ ”»…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group11", tvwChild, "Report38", "  ﬁ—Ì— »√—’œ… Ã„Ì⁄ «·⁄„·«¡ Ê«·„Ê—œÌ‰", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group12", " ﬁ«—Ì— «·Œ“‰", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group12", tvwChild, "Report39", " ﬁ—Ì— »ﬂ‘› Õ”«» «·Œ“‰…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group12", tvwChild, "Report40", " ﬁ—Ì— »⁄„·Ì«  «·”Õ» Ê«·≈Ìœ«⁄ „‰ «·Œ“‰…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group12", tvwChild, "Report41", " ﬁ—Ì— »«·“Ì«œ… Ê«·⁄Ã— ›Ï ‰ﬁœÌ… «·Œ“‰…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group12", tvwChild, "Report42", " ﬁ—Ì— »«·Ã—œ «·ÌÊ„Ï  ··Œ“‰…", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Root", tvwChild, "Group13", "«· ﬁ«—Ì— «·≈Ã„«·Ì« ", "Close", "Root")
        XNode.ExpandedImage = "OpenFolder"
        Set XNode = .Nodes.Add("Group13", tvwChild, "Report43", " ﬁ—Ì— «·‘Â—", "Purchase", "Purchase")
        Set XNode = .Nodes.Add("Group13", tvwChild, "Report46", " ﬁ«—Ì— «·√—»«Õ", "Purchase", "Purchase")
        .Nodes("Root").Expanded = True
        .Nodes("Root").EnsureVisible

        For i = 1 To .Nodes.count

            If .Nodes(i).key Like "Group*" Then
                .Nodes(i).ForeColor = &H80&
            End If

        Next i

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Exit Sub
ErrTrap:
End Sub

Private Sub TrvReports_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim Msg As String
    Dim i As Integer

    'Me.TrvReports.Enabled = False
    On Error Resume Next

    If Node.key Like "Group*" Then
        Exit Sub
    End If

    WriteCaption lbl(3), Node.text
    HideFraOptions

    Select Case Node.key

        Case Is = "Group*"

        Case "Report1"
            Msg = "⁄—÷  ﬁ—Ì— »«·„»Ì⁄«  «·‰ﬁœÌ… «Ê «·√Ã·… «Ê ﬂ·«Â„« ( ⁄—÷ «·›Ê« Ì—) «· Ï ÕœÀ  ›Ï «·‘—ﬂ… ..ÊÌ„ﬂ‰ﬂ ≈Œ Ì«—"
            Msg = Msg & Chr(13) & "Œ“‰… „⁄Ì‰…  „  ”ÃÌ· ﬁÌ„ «·›Ê« Ì— ›ÌÂ«"
            WriteCaption lbl(4), Msg
            Me.TxtNodeReport.text = "Report1"

            With Me.CboSortData
                .Clear
                .AddItem "—ﬁ„  ”ÃÌ· «·›« Ê—… ›Ï «·»—‰«„Ã"
                .AddItem "—ﬁ„ „”·”· «·›« Ê—… ›Ï «·»—‰«„Ã"
                .AddItem " «—ÌŒ «·›« Ê—…"
                .AddItem "«”„ «·Œ“‰…"
                .AddItem "«”„ «·„Œ“‰"
                .AddItem "«”„ «·⁄„Ì·"
                .AddItem "«”„ «·„ÊŸ›"
                .AddItem "ﬁÌ„… «·›« Ê—… «·√Ã„«·Ì…"
            End With

            ShowTab 0

        Case "Report27"
            '-------- ﬁ«—Ì— «·„’—Ê›« 
            Msg = " ’„Ì„  ﬁ«—Ì— «·„’—Ê›«  »ÕÌÀ Ì„ﬂ‰ﬂ ≈Œ Ì«— ‰Ê⁄ «·„’—Ê›«  «·–Ï  —ÌœÂ "
            Msg = Msg & Chr(13) & "√Ê ⁄—÷ ⁄„·Ì«  ’—› „Õœœ… „‰ —ﬁ„ ≈·Ï —ﬁ„ -«Ê  ÕœÌœ ﬁÌ„… „»·€ „Õœœ "
            Msg = Msg & Chr(13) & "Ê«”„ «·Œ“‰… «·„’—Ê› „‰Â«...Ê√Ì÷« Ì„ﬂ‰ﬂ  ÕœÌœ «”„ «·„” Œœ„ «·„”Ã· ·⁄„·Ì… «·’—›"
            WriteCaption lbl(4), Msg
            Me.TxtNodeReport.text = "Report27"

            With Me.CboSortData
                .Clear
                .AddItem "«”„ ‰Ê⁄ «·„’—Ê›« "
                .AddItem "—ﬁ„ ⁄„·Ì… «·’—›"
                .AddItem " «—ÌŒ ⁄„·Ì… «·’—›"
                .AddItem "«”„ «·Œ“‰…"
                .AddItem "÷ ﬁÌ„… „»·€ «·’—›"
                .ListIndex = 0
            End With

            With Me.CboReportStyle
                .Clear
                .AddItem "⁄—÷ „  «·Ï ·⁄„·Ì«  «·’—›"
                .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» ‰Ê⁄ «·’—›"
                .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» «”„ «·Œ“‰…"
                .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» «”„ «·„” Œœ„ «·„Õ——"
                .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» ﬁÌ„… „»·€ «·’—›"
                .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”»  «—ÌŒ «·’—›"
                .ListIndex = 0
            End With

            ShowTab 2

        Case "Report28"
            Msg = " ’„Ì„  ﬁ«—Ì— «·„œ›Ê⁄«  »ÕÌÀ Ì„ﬂ‰ﬂ ≈Œ Ì«— «”„ «·⁄„Ì· «Ê «·„Ê—œ «·–Ï  —ÌœÂ "
            Msg = Msg & Chr(13) & "√Ê ⁄—÷ ⁄„·Ì«  „œ›Ê⁄«  „Õœœ… „‰ —ﬁ„ ≈·Ï —ﬁ„ -«Ê  ÕœÌœ ﬁÌ„… „»·€ „Õœœ "
            Msg = Msg & Chr(13) & "Ê«”„ «·Œ“‰… «·„œ›Ê⁄ „‰Â«...Ê√Ì÷« Ì„ﬂ‰ﬂ  ÕœÌœ «”„ «·„” Œœ„ «·„”Ã· ·⁄„·Ì… «·„œ›Ê⁄« "
            WriteCaption lbl(4), Msg
            Me.TxtNodeReport.text = "Report28"

            With Me.CboSortData
                .Clear
                .AddItem "«”„ «·⁄„Ì· «Ê «·„Ê—œ"
                .AddItem "—ﬁ„ ⁄„·Ì… «·„œ›Ê⁄« "
                .AddItem " «—ÌŒ ⁄„·Ì… «·„œ›Ê⁄« "
                .AddItem "«”„ «·Œ“‰…"
                .AddItem "ﬁÌ„… „»·€ «·„œ›Ê⁄« "
            End With

            FraReportOptions(0).Visible = True

        Case "Report44"
            Msg = "Ì⁄—÷ ·ﬂ Â–« «· ﬁ—Ì— «·√’‰«› «· Ï —’ÌœÂ« «·Õ«·Ï ’›— ›Ï «·»—‰«„Ã"
            WriteCaption lbl(4), Msg
            Me.TxtNodeReport.text = "Report44"

            With Me.CboSortData
                .Clear
                .AddItem "—ﬁ„ «·’‰›"
                .AddItem "ﬂÊœ «·’‰›"
                .AddItem "«”„ «·’‰›"
                .AddItem "—’Ìœ «·’‰›"
            End With

            With Me.CboReportStyle
                .Clear
                .AddItem " Ã„Ì⁄ «·»Ì«‰«  »‰«¡ ⁄·Ï „Ã„Ê⁄«  «·√’‰«›"
                .AddItem "⁄—÷ ÃœÊ· ·√”„«¡ «·√’‰«›"
                .ListIndex = 0
            End With

            ShowTab 1
       
        Case "Report45"
            Msg = "Ì⁄—÷ Â–« «· ﬁ—Ì—  ›«’Ì· ⁄„·Ì«   Ã«—Ì… „⁄‰Ì… ⁄·Ï „Ã„Ê⁄… √’‰«› ﬂ«„·…"
            WriteCaption lbl(4), Msg
            Me.TxtNodeReport.text = "Report45"

            With Me.CboSortData
                .Clear
                .AddItem "—ﬁ„ «·’‰›"
                .AddItem "ﬂÊœ «·’‰›"
                .AddItem "«”„ «·’‰›"
                .AddItem "≈Ã„«·Ï ÕÃ„ «·⁄„·Ì…"
            End With

            With Me.CboReportStyle
                .Clear
                .AddItem "„Ã„Ê⁄«  «·«’‰«› „— »… Õ”» «”„ «·„Ã„Ê⁄…"
                .AddItem "„Ã„Ê⁄«  «·«’‰«› „— »… Õ”» ﬂÊœ «·„Ã„Ê⁄…"
                .ListIndex = 0
            End With

            ShowTab 4

        Case "Report46"
            Msg = " ﬁ—Ì— «·√—»«Õ „‰ Œ·«· ›Ê« Ì— «·„»Ì⁄«  ÌﬁÊ„ «·»—‰«„Ã »Õ”«» «·—»Õ „‰ ›—ﬁ "
            Msg = Msg & "”⁄— «· ﬂ·›… ›Ï «·›« Ê—… Ê”⁄— «·»Ì⁄ "
            WriteCaption lbl(4), Msg
            Me.TxtNodeReport.text = "Report46"

            With Me.CboReportStyle
                .Clear
            End With

            ShowTab 5
        
        Case Else
            Me.lbl(3).Caption = ""
            Me.lbl(4).Caption = ""
    End Select

    Me.TrvReports.SetFocus
End Sub

Private Sub HideFraOptions()
    Dim i As Integer
    On Error Resume Next

    'For I = 0 To TabMain.NumTabs - 1
    '    TabMain.TabVisible(I) = False
    'Next I
    TabMain.Visible = False
    Me.Fra(0).Visible = False
    Me.Fra(3).Visible = False
    Me.Fra(4).Visible = False
End Sub

Private Sub ShowExpensesReport(m_PrintTarget As PrintTarget)

    Dim StrSQL As String
    Dim Msg As String
    Dim StrDesReport As String
    Dim BolBegine As Boolean
    Dim cReport As ClsRepoerts
    Dim StrOrderBy  As String
    Dim StrSortGroup As String
    Dim StrSortData As String
    Dim IntReportStyle As Integer

    On Error GoTo ErrTrap

    StrSQL = "Select * From ExpensesReport"
    BolBegine = False

    If DCboExpensesName.BoundText <> "" Then
        StrDesReport = StrDesReport & Chr(13) & "‰Ê⁄ «·„’—Ê›«  :- " & Me.DCboExpensesName.text

        If BolBegine = True Then
            StrSQL = StrSQL + " and ExpensesID=" & DCboExpensesName.BoundText & ""
        Else
            StrSQL = StrSQL + " where ExpensesID=" & DCboExpensesName.BoundText & ""
            BolBegine = True
        End If
    End If

    If val(Me.TxtFromID(2).text) > 0 Then
        StrDesReport = StrDesReport & Chr(13) & "«—ﬁ«„ «·Õ—ﬂ«   »œ« „‰ : " & val(Me.TxtFromID(2).text)

        If BolBegine = True Then
            StrSQL = StrSQL + " AND NoteSerial >=" & val(Me.TxtFromID(2).text) & ""
        Else
            StrSQL = StrSQL + " WHERE NoteSerial >=" & val(Me.TxtFromID(2).text) & ""
            BolBegine = True
        End If
    End If

    If val(Me.TxtToID(2).text) > 0 Then
        StrDesReport = StrDesReport & Chr(13) & "«—ﬁ«„ «·Õ—ﬂ«   ‰ ÂÏ Õ Ï : " & val(Me.TxtToID(2).text)

        If BolBegine = True Then
            StrSQL = StrSQL + " AND NoteSerial <=" & val(Me.TxtToID(2).text) & ""
        Else
            StrSQL = StrSQL + " WHERE NoteSerial <=" & val(Me.TxtToID(2).text) & ""
            BolBegine = True
        End If
    End If

    If Me.DcboBoxTrans.BoundText <> "" Then
        StrDesReport = StrDesReport & Chr(13) & "Ã—  ⁄·Ï «·Œ“‰…: " & Me.DcboBoxTrans.text

        If BolBegine = True Then
            StrSQL = StrSQL + " and BoxID = " & Me.DcboBoxTrans.BoundText & ""
        Else
            StrSQL = StrSQL + " Where BoxID =" & Me.DcboBoxTrans.BoundText & ""
            BolBegine = True
        End If
    End If

    If val(Me.TxtValue(1).text) > 0 Then
        If Me.Opt(7).value = True Then
            If BolBegine = True Then
                StrSQL = StrSQL + " AND Note_Value =" & val(Me.TxtValue(1).text) & ""
            Else
                StrSQL = StrSQL + " Where Note_Value =" & val(Me.TxtValue(1).text) & ""
                BolBegine = True
            End If

            StrDesReport = StrDesReport & Chr(13) & "ﬁÌ„… «·„’—Ê›«   ”«ÊÏ:" & val(Me.TxtValue(1).text)
        ElseIf Me.Opt(6).value = True Then

            If BolBegine = True Then
                StrSQL = StrSQL + " AND Note_Value >" & val(Me.TxtValue(1).text) & ""
            Else
                StrSQL = StrSQL + " where Note_Value >" & val(Me.TxtValue(1).text) & ""
                BolBegine = True
            End If

            StrDesReport = StrDesReport & Chr(13) & "ﬁÌ„… «·„’—Ê›«  «ﬂ»— „‰ :" & val(Me.TxtValue(1).text)
        Else

            If BolBegine = True Then
                StrSQL = StrSQL + " AND Note_Value <" & val(Me.TxtValue(1).text) & ""
            Else
                StrSQL = StrSQL + " Where Note_Value <" & val(Me.TxtValue(1).text) & ""
                BolBegine = True
            End If

            StrDesReport = StrDesReport & Chr(13) & "ﬁÌ„… «·„’—Ê›«  √ﬁ· „‰:" & val(Me.TxtValue(1).text)
        End If
    End If

    If Me.DcboUsers1.BoundText <> "" Then
        StrDesReport = StrDesReport & Chr(13) & "”Ã·Â« «·„” Œœ„:" & Me.DcboUsers1.text

        If BolBegine = True Then
            StrSQL = StrSQL + " and UserID= " & Me.DcboUsers1.BoundText & ""
        Else
            StrSQL = StrSQL + " Where UserID=" & Me.DcboUsers1.BoundText & ""
            BolBegine = True
        End If
    End If

    If Not IsNull(DTPFrom.value) Then
        StrDesReport = StrDesReport & Chr(13) & "»œ«Ì…  ”ÃÌ· «·Õ—ﬂ« :" & DisplayDate(DTPFrom.value)

        If BolBegine = True Then
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND NoteDate >=#" & SQLDate(DTPFrom.value) & "#"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND NoteDate >='" & SQLDate(DTPFrom.value) & "'"
            End If

        Else

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " WHERE NoteDate >=#" & SQLDate(DTPFrom.value) & "#"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " WHERE NoteDate >='" & SQLDate(DTPFrom.value) & "'"
            End If

            BolBegine = True
        End If
    End If

    If Not IsNull(DTPTo.value) Then
        StrDesReport = StrDesReport & Chr(13) & "‰Â«Ì…  ”ÃÌ· «·Õ—ﬂ« :" & DisplayDate(DTPTo.value)

        If BolBegine = True Then
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " and NoteDate <=#" & SQLDate(DTPTo.value) & "#"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " and NoteDate <='" & SQLDate(DTPTo.value) & "'"
            End If

        Else

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " where NoteDate <=#" & SQLDate(DTPTo.value) & "#"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " where NoteDate <='" & SQLDate(DTPTo.value) & "'"
            End If
        End If
    End If

    '--------------------------------------------
    'With Me.CboSortData
    '    .Clear
    '    .AddItem "«”„ ‰Ê⁄ «·„’—Ê›« "
    '    .AddItem "—ﬁ„ ⁄„·Ì… «·’—›"
    '    .AddItem " «—ÌŒ ⁄„·Ì… «·’—›"
    '    .AddItem "«”„ «·Œ“‰…"
    '    .AddItem "ﬁÌ„… „»·€ «·’—›"
    '    .ListIndex = 0
    'End With
    'With Me.CboReportStyle
    '    .Clear
    '    .AddItem "⁄—÷ „  «·Ï ·⁄„·Ì«  «·’—›"
    '    .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» ‰Ê⁄ «·’—›"
    '    .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» «”„ «·Œ“‰…"
    '    .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» «”„ «·„” Œœ„ «·„Õ——"
    '    .AddItem " Ã„Ì⁄ «·»Ì«‰«  Õ”» ﬁÌ„… „»·€ «·’—›"
    '    .ListIndex = 0
    'End With
    StrOrderBy = ""

    If Me.CboReportStyle.ListIndex = 0 Then
        StrSortGroup = "ExpensesReport.Name"
        IntReportStyle = 1
    ElseIf Me.CboReportStyle.ListIndex = 1 Then
        StrSortGroup = "ExpensesReport.Name"
        IntReportStyle = 1
    ElseIf Me.CboReportStyle.ListIndex = 2 Then
        StrSortGroup = "ExpensesReport.BoxName"
        IntReportStyle = 2
    ElseIf Me.CboReportStyle.ListIndex = 3 Then
        StrSortGroup = "ExpensesReport.UserName"
        IntReportStyle = 3
    ElseIf Me.CboReportStyle.ListIndex = 4 Then
        StrSortGroup = "ExpensesReport.Note_Value"
        IntReportStyle = 4
    End If

    '--------------------------------------------
    If Me.CboSortData.ListIndex = 0 Then
        StrSortData = "ExpensesReport.Name"
    ElseIf Me.CboSortData.ListIndex = 1 Then
        StrSortData = "ExpensesReport.NoteSerial"
    ElseIf Me.CboSortData.ListIndex = 2 Then
        StrSortData = "ExpensesReport.NoteDate"
    ElseIf Me.CboSortData.ListIndex = 3 Then
        StrSortData = "ExpensesReport.BoxName"
    ElseIf Me.CboSortData.ListIndex = 4 Then
        StrSortData = "ExpensesReport.Note_Value"
    End If

    '--------------------------------------------
    If (StrSortGroup = StrSortData) And StrSortData <> "" Then
        'The (Sort Group) Is equal to (Sort Data)
        StrSortData = ""
    End If

    If Trim$(StrSortGroup) <> "" Then
        If OptSort(3).value = True Then
            StrSortGroup = StrSortGroup + " ASC "
        ElseIf OptSort(2).value = True Then
            StrSortGroup = StrSortGroup + " DESC"
        End If
    End If

    If Trim$(StrSortData) <> "" Then
        If OptSort(0).value = True Then
            StrSortData = StrSortData + " ASC "
        ElseIf OptSort(1).value = True Then
            StrSortData = StrSortData + " DESC"
        End If
    End If

    '--------------------------------------------
    If StrSortGroup <> "" Then
        StrOrderBy = " Order BY " & StrSortGroup & IIf(StrSortData <> "", ",", "")
    Else
        StrOrderBy = " Order BY " & StrSortData
    End If

    StrSQL = StrSQL + StrOrderBy
    '--------------------------------------------
    Set cReport = New ClsRepoerts
    cReport.ExpensesReports StrSQL, m_PrintTarget, StrDesReport, CBool(Me.Chk(5).value), IntReportStyle, CBool(Me.ChkChart.value), IIf(Me.OptChart(0).value = True, 0, 1)
    Exit Sub
ErrTrap:
End Sub

Private Sub adjustFraToShow(FraDueOptions As Frame)

    With FraDueOptions
        .left = 30
        .top = 1650
    End With

End Sub

Private Sub ShowZeroItemsStock(m_PrintTarget As PrintTarget)
    Dim StrSQL  As String
    Dim cReport As ClsItemsReport
    Dim IntReportStyle As Integer
    Dim StrOrderBy  As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT TOP 100 PERCENT QryGARDShort.QTY, QryGARDShort.ItemID," & "QryGARDShort.ItemCode, QryGARDShort.ItemName, QryGARDShort.StoreID," & "QryGARDShort.StoreName , QryGARDShort.GroupID, dbo.Groups.GroupName "
        StrSQL = StrSQL + " FROM dbo.QryGARDShort() QryGARDShort INNER JOIN " & "dbo.Groups ON QryGARDShort.GroupID = dbo.Groups.GroupID"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        Exit Sub
    End If

    '---------------------------------------------------------
    If Opt(3).value = True Then
        StrSQL = StrSQL + " Where (QTY = 0)"
    ElseIf Opt(4).value = True Then
        StrSQL = StrSQL + " Where (QTY < 0)"
    ElseIf Opt(5).value = True Then
        StrSQL = StrSQL + " Where (QTY <= 0)"
    End If

    '---------------------------------------------------------
    If Opt(11).value = True Then
        '—’Ìœ «·’‰› ’›— ·«‰Â ·„  ”Ã· ⁄·ÌÂ «Ï Õ—ﬂ«   Ã«—Ì…
        StrSQL = StrSQL + " AND (ItemID NOT IN "
        StrSQL = StrSQL + "(SELECT DISTINCT TOP 100 PERCENT Item_ID "
        StrSQL = StrSQL + " FROM dbo.Transaction_Details "
        StrSQL = StrSQL + " ORDER BY Item_ID))"
    ElseIf Opt(10).value = True Then
        '—’Ìœ «·’‰› ’›— ·«‰Â  ≈‰ ÂÏ „‰ «·„Œ“‰
        StrSQL = StrSQL + " AND (ItemID IN "
        StrSQL = StrSQL + "(SELECT DISTINCT TOP 100 PERCENT Item_ID "
        StrSQL = StrSQL + " FROM dbo.Transaction_Details "
        StrSQL = StrSQL + " ORDER BY Item_ID))"
    ElseIf Opt(9).value = True Then
        'ﬂ· «·Õ«·«  ...·”  „Õ «Ã ≈·Ï «Ï ‘—Êÿ
    End If

    '---------------------------------------------------------
    StrOrderBy = " Order By "

    If CboReportStyle.ListIndex = 0 Then
        IntReportStyle = 0
        StrOrderBy = StrOrderBy + "GroupName"
    ElseIf CboReportStyle.ListIndex = 1 Then
        IntReportStyle = 1
        StrOrderBy = StrOrderBy + "QryGARDShort.GroupID"
    End If

    If Me.OptSort(2).value = True Then
        StrOrderBy = StrOrderBy + " DESC"
    ElseIf Me.OptSort(3).value = True Then
        StrOrderBy = StrOrderBy + " ASC"
    End If

    If Me.CboSortData.ListIndex > 0 Then
        If Me.CboSortData.ListIndex = 0 Then
            StrOrderBy = StrOrderBy + ",ItemID"
        ElseIf Me.CboSortData.ListIndex = 1 Then
            StrOrderBy = StrOrderBy + ",ItemCode"
        ElseIf Me.CboSortData.ListIndex = 2 Then
            StrOrderBy = StrOrderBy + ",ItemName"
        ElseIf Me.CboSortData.ListIndex = 3 Then
            StrOrderBy = StrOrderBy + ",Qty"
        End If

    Else
        StrOrderBy = StrOrderBy + ",ItemID"
    End If

    If Me.OptSort(0).value = True Then
        StrOrderBy = StrOrderBy + " ASC"
    ElseIf Me.OptSort(1).value = True Then
        StrOrderBy = StrOrderBy + " DESC"
    End If

    StrSQL = StrSQL + StrOrderBy
    Set cReport = New ClsItemsReport
    cReport.ShowZeroItemsStock StrSQL, m_PrintTarget, IntReportStyle
End Sub

Private Sub ShowSalesReport(m_PrintTarget As PrintTarget)
    Dim StrSQL As String
    Dim StrOrderBy As String
    Dim Reports As ClsRepoerts

    StrSQL = "Select * From ReportSallingTime "
    StrSQL = StrSQL + " Where Transaction_ID <> 0"

    If Me.CboPaymentMethod.ListIndex = 0 Then
        StrSQL = StrSQL + " AND PaymentType=0"
    ElseIf Me.CboPaymentMethod.ListIndex = 1 Then
        StrSQL = StrSQL + " AND PaymentType=1"
    End If

    If val(Me.DcboBoxTrans.BoundText) <> 0 Then
        StrSQL = StrSQL + " AND BoxID=" & val(Me.DcboBoxTrans.BoundText)
    End If

    If val(Me.DcboStores.BoundText) <> 0 Then
        StrSQL = StrSQL + " AND StoreID=" & val(Me.DcboStores.BoundText)
    End If

    If val(Me.DcboEmpID.BoundText) <> 0 Then
        StrSQL = StrSQL + " AND Emp_ID=" & val(Me.DcboEmpID.BoundText)
    End If

    '
    If val(Me.DcboCustomersSuppliers1.BoundText) <> 0 Then
        StrSQL = StrSQL + " AND CusID=" & val(Me.DcboCustomersSuppliers1.BoundText)
    End If

    If Me.CboSalesType.ListIndex >= 0 Then
        If Me.CboSalesType.ListIndex = 0 Then
            StrSQL = StrSQL + " AND SaleType=0"
        ElseIf Me.CboSalesType.ListIndex = 1 Then
            StrSQL = StrSQL + " AND SaleType=1"
        End If
    End If

    If val(Me.Txt(0).text) > 0 Then
        StrSQL = StrSQL + " AND TotalAfterTax=" & val(Me.Txt(0).text)
    End If

    If Not IsNull(DTPFrom.value) Then
        StrSQL = StrSQL + " and Transaction_Date >=" & SQLDate(DTPFrom.value, True) & ""
    End If

    If Not IsNull(DTPTo.value) Then
        StrSQL = StrSQL + " and Transaction_Date <=" & SQLDate(DTPTo.value, True) & ""
    End If

    If Me.CboSortData.ListIndex >= 0 Then
        StrOrderBy = " Order By "

        If Me.CboSortData.ListIndex = 0 Then
            StrOrderBy = StrOrderBy + "Transaction_ID"
        ElseIf Me.CboSortData.ListIndex = 1 Then
            StrOrderBy = StrOrderBy + "Transaction_Serial"
        ElseIf Me.CboSortData.ListIndex = 2 Then
            StrOrderBy = StrOrderBy + "Transaction_Date"
        ElseIf Me.CboSortData.ListIndex = 3 Then
            StrOrderBy = StrOrderBy + "BoxName"
        ElseIf Me.CboSortData.ListIndex = 4 Then
            StrOrderBy = StrOrderBy + "StoreName"
        ElseIf Me.CboSortData.ListIndex = 5 Then
            StrOrderBy = StrOrderBy + "CusName"
        ElseIf Me.CboSortData.ListIndex = 6 Then
            StrOrderBy = StrOrderBy + "Emp_Name"
        ElseIf Me.CboSortData.ListIndex = 7 Then
            StrOrderBy = StrOrderBy + "TotalAfterTax"
        End If

        If OptSort(0).value = True Then
            StrOrderBy = StrOrderBy + " ASC "
        ElseIf OptSort(1).value = True Then
            StrOrderBy = StrOrderBy + " DESC "
        End If
    End If

    StrSQL = StrSQL + StrOrderBy
    Set Reports = New ClsRepoerts
    Reports.ShowSallingTime StrSQL, Me.DTPFrom.value, Me.DTPTo.value, False
End Sub

Private Sub ShowTab(IntTabIndex As Integer)
    Me.TabMain.Visible = True
    Me.TabMain.CurrTab = IntTabIndex
    Me.Fra(0).Visible = True
    Me.Fra(3).Visible = True
    Me.Fra(4).Visible = True
End Sub

Private Sub ShowGroupsItemsTransactions(m_PrintTarget As PrintTarget)
    Dim StrSQL As String
    Dim Msg As String
    Dim IntTransType As Integer
    Dim cReport As ClsItemsReport

    If CboTransactions.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·Õ—ﬂ… ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    ElseIf Me.CboTransactions.ListIndex = 0 Then
        IntTransType = 2
    ElseIf Me.CboTransactions.ListIndex = 1 Then
        IntTransType = 1
    ElseIf Me.CboTransactions.ListIndex = 2 Then
        IntTransType = 9
    ElseIf Me.CboTransactions.ListIndex = 3 Then
        IntTransType = 5
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryItemsTransactionsTotals.ItemID,QryItemsTransactionsTotals.ItemCode," & "QryItemsTransactionsTotals.ItemName,"
        StrSQL = StrSQL + "QryItemsTransactionsTotals.GroupID,dbo.Groups.GroupName," & "QryItemsTransactionsTotals.Total," & "QryItemsTransactionsTotals.TotalQty "
    
        StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(" & IntTransType & ",0,"

        If IsNull(Me.DTPFrom.value) Then
            StrSQL = StrSQL + SQLDate(CDate(#1/1/1900#), True)
        Else
            StrSQL = StrSQL + SQLDate(Me.DTPFrom.value, True)
        End If

        StrSQL = StrSQL + ","

        If IsNull(Me.DTPTo.value) Then
            StrSQL = StrSQL + SQLDate(CDate(#1/1/9999#), True)
        Else
            StrSQL = StrSQL + SQLDate(Me.DTPTo.value, True)
        End If

        StrSQL = StrSQL + ")"
        StrSQL = StrSQL + " QryItemsTransactionsTotals INNER JOIN dbo.Groups ON QryItemsTransactionsTotals.GroupID = dbo.Groups.GroupID "

        If val(Me.DcboGroups1.BoundText) <> 0 Then
            StrSQL = StrSQL + " Where QryItemsTransactionsTotals.GroupID=" & val(Me.DcboGroups1.BoundText)
        End If

        StrSQL = StrSQL + " Order By "

        If Me.CboReportStyle.ListIndex = -1 Or Me.CboReportStyle.ListIndex = 0 Then
            StrSQL = StrSQL + " GroupName "
        Else
            StrSQL = StrSQL + " GroupID "
        End If

        If OptSort(3).value = True Then
            StrSQL = StrSQL + " ASC "
        ElseIf OptSort(3).value = True Then
            StrSQL = StrSQL + " DESC "
        End If

        StrSQL = StrSQL + ","

        If Me.CboSortData.ListIndex = -1 Or Me.CboSortData.ListIndex = 0 Then
            StrSQL = StrSQL + "ItemID"
        ElseIf Me.CboSortData.ListIndex = 1 Then
            StrSQL = StrSQL + "ItemCode"
        ElseIf Me.CboSortData.ListIndex = 2 Then
            StrSQL = StrSQL + "ItemName"
        ElseIf Me.CboSortData.ListIndex = 3 Then
            StrSQL = StrSQL + "Total"
        End If

        If OptSort(0).value = True Then
            StrSQL = StrSQL + " ASC "
        ElseIf OptSort(1).value = True Then
            StrSQL = StrSQL + " DESC "
        End If
    
    Else
    End If

    Set cReport = New ClsItemsReport
    cReport.ShowGroupsItemsTrans StrSQL, m_PrintTarget, IntTransType, Me.DTPFrom.value, Me.DTPTo.value
        
End Sub

Private Sub ShowProfits(m_PrintTarget As PrintTarget)

End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Trim$(Me.TxtItemCode.text) <> "" Then
            Me.DcboItemName.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        Else
            Me.DcboItemName.BoundText = ""
        End If
    End If

End Sub

