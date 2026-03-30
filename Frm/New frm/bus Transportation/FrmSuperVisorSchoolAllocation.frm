VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSuperVisorSchoolAllocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   ⁄„·Ì«   Œ’Ì’ „‘—›Ì‰ ··„œ«—”"
   ClientHeight    =   8880
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmSuperVisorSchoolAllocation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   13440
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   2076
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8880
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13440
      _cx             =   23707
      _cy             =   15663
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   732
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   720
         Width           =   13212
         _cx             =   23310
         _cy             =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   10320
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1572
         End
         Begin MSComCtl2.DTPicker Date 
            Height          =   312
            Left            =   1308
            TabIndex        =   43
            Top             =   240
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98762753
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal DateH 
            Height          =   252
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1116
            _ExtentX        =   1958
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   6960
            TabIndex        =   46
            Top             =   240
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   315
            Left            =   3840
            TabIndex        =   49
            Top             =   240
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   17
            Left            =   5805
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   11
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   768
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· Œ’Ì’ "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   8
            Left            =   2556
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Œ’Ì’"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   7
            Left            =   11880
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1092
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1560
         Width           =   13212
         _cx             =   23310
         _cy             =   1931
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox txtMinistry 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   1560
         End
         Begin VB.TextBox txtemp_code 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10332
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   1560
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   312
            Left            =   4476
            TabIndex        =   3
            Top             =   240
            Width           =   1248
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98762753
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   312
            Left            =   4476
            TabIndex        =   7
            Top             =   600
            Width           =   1248
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98762753
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcEmp 
            Height          =   315
            Left            =   6945
            TabIndex        =   6
            Top             =   600
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   456
            Index           =   9
            Left            =   2160
            TabIndex        =   9
            Top             =   360
            Width           =   876
            _ExtentX        =   1535
            _ExtentY        =   794
            ButtonPositionImage=   1
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":038A
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
         Begin MSDataListLib.DataCombo dcMA 
            Height          =   315
            Left            =   6960
            TabIndex        =   2
            Top             =   240
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin Dynamic_Byte.NourHijriCal FromDateH 
            Height          =   252
            Left            =   3180
            TabIndex        =   4
            Top             =   240
            Width           =   1224
            _ExtentX        =   2170
            _ExtentY        =   450
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   252
            Left            =   3180
            TabIndex        =   8
            Top             =   600
            Width           =   1224
            _ExtentX        =   2170
            _ExtentY        =   450
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   456
            Index           =   5
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   864
            _ExtentX        =   1535
            _ExtentY        =   794
            ButtonPositionImage=   1
            Caption         =   "Õ–› ”ÿ—"
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":6BEC
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
            Height          =   456
            Index           =   8
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Width           =   864
            _ExtentX        =   1535
            _ExtentY        =   794
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ·"
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":D44E
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·Ê“«—Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   1
            Left            =   11952
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   1116
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·»œ√ "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   9
            Left            =   5712
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   1068
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«‰ Â«¡  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   10
            Left            =   5604
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—›"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   6
            Left            =   9180
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   600
            Width           =   888
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„‘—›"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   2
            Left            =   11964
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   600
            Width           =   1116
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œ—”…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   5
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   684
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   4512
         Left            =   0
         TabIndex        =   11
         Top             =   2760
         Width           =   13380
         _cx             =   23601
         _cy             =   7959
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSuperVisorSchoolAllocation.frx":13CB0
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
         ExplorerBar     =   1
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   648
         Index           =   5
         Left            =   0
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   13512
         _cx             =   23839
         _cy             =   1138
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
         Picture         =   "FrmSuperVisorSchoolAllocation.frx":13EA4
         Caption         =   "    ⁄„·Ì«   Œ’Ì’ „‘—›Ì‰ ··„œ«—”   "
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
         PicturePos      =   0
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
         Begin VB.TextBox txtIDSA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   21
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":14B7E
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
            Height          =   375
            Index           =   2
            Left            =   630
            TabIndex        =   22
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":14F18
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
            Height          =   375
            Index           =   1
            Left            =   2220
            TabIndex        =   23
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":152B2
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
            Height          =   375
            Index           =   3
            Left            =   1155
            TabIndex        =   24
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":1564C
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
         Height          =   492
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   7440
         Width           =   5772
         _cx             =   10186
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   312
            Index           =   4
            Left            =   816
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   120
            Width           =   1104
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   312
            Index           =   0
            Left            =   3816
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   120
            Width           =   1104
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   120
            Width           =   660
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   2940
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   828
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   696
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   8040
         Width           =   13212
         _cx             =   23310
         _cy             =   1217
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
            Height          =   456
            Index           =   0
            Left            =   11508
            TabIndex        =   31
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":159E6
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
            Height          =   456
            Index           =   1
            Left            =   9816
            TabIndex        =   32
            Top             =   120
            Width           =   1644
            _ExtentX        =   2910
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":1C248
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
            Height          =   456
            Index           =   2
            Left            =   8196
            TabIndex        =   33
            Top             =   120
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":22AAA
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
            Height          =   456
            Index           =   3
            Left            =   6552
            TabIndex        =   34
            Top             =   120
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":2930C
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
            Height          =   456
            Index           =   4
            Left            =   5004
            TabIndex        =   35
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":2FB6E
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
            Height          =   456
            Index           =   6
            Left            =   1848
            TabIndex        =   36
            Top             =   120
            Width           =   1644
            _ExtentX        =   2910
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":363D0
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
            Height          =   456
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":5FFF2
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
            Height          =   456
            Index           =   7
            Left            =   3504
            TabIndex        =   38
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   794
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
            ButtonImage     =   "FrmSuperVisorSchoolAllocation.frx":66854
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
   End
End
Attribute VB_Name = "FrmSuperVisorSchoolAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp2  As ADODB.Recordset
Dim TTP As clstooltip

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            txtID.Text = CStr(new_id("TblSupervisorAllocation", "IDSA", "", True))
         '   txtName.SetFocus
            Grid.Rows = 1
        Case 1
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "E"
        Case 2
            SaveData
        Case 3
            Undo
        Case 4
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_Company
        Case 5
                DeleteRow
        Case 6
            Unload Me
         Case 7
            'print_report2
            Case 8
            Grid.Rows = 1
         Case 9
            addrow
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub DeleteRow()

If Grid.Row < Grid.FixedRows Then Exit Sub:

Dim i As Integer
i = Grid.Row

Grid.RemoveItem (i)

End Sub


Private Sub addrow()

If dcMA.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox ("„‰ ›÷··ﬂ «Œ — «·„œ—”…")
        Else
                MsgBox ("Select School")
        End If
        dcMA.SetFocus
        Exit Sub
End If


If dcEmp.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«Œ — «·„‘—›  «Ê·«")
        Else
        MsgBox ("Select Supervisor ")
        End If
        dcEmp.SetFocus
        SendKeys ("{F4}")
        Exit Sub
End If

Dim l As Integer


With Grid
For l = 1 To .Rows - 1
        If .TextMatrix(l, .ColIndex("schoolFileID")) = dcMA.BoundText Then
                MsgBox (" „  Œ’Ì’ „‘—› ·Â–Â «·œ—”…")
                Exit Sub
        End If
Next


     Dim str As String
     str = " Select * from TblSupervisorAllocation h,  TblSupervisorAllocation_details d"
     str = str & "  where  h.IDSA = d.idsa and  h.DurationID = " & val(dcDuration.BoundText) & " and d.SchoolFileID  = " & val(dcMA.BoundText)
     Set Rs_Temp = New ADODB.Recordset
     Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
      If Rs_Temp.RecordCount > 0 Then
                 MsgBox (" „  Œ’Ì’ „‘—› ·Â–Â «·œ—”…")
                Exit Sub
      End If



Dim i As Integer
Grid.Rows = Grid.Rows + 1
i = Grid.Rows
i = i - 1

  .TextMatrix(i, .ColIndex("Serial")) = i
  .TextMatrix(i, .ColIndex("schoolFileID")) = dcMA.BoundText
  .TextMatrix(i, .ColIndex("SchoolFileName")) = dcMA.Text
  .TextMatrix(i, .ColIndex("emp_id")) = dcEmp.BoundText
  .TextMatrix(i, .ColIndex("emp_name")) = dcEmp.Text
  .TextMatrix(i, .ColIndex("FromDate")) = FromDate.value
  .TextMatrix(i, .ColIndex("FromDateH")) = FromDateH.value
  .TextMatrix(i, .ColIndex("ToDate")) = ToDate.value
  .TextMatrix(i, .ColIndex("ToDateH")) = ToDateH.value
 .TextMatrix(i, .ColIndex("ministryno")) = txtMinistry.Text
  
End With
'

dcEmp.BoundText = ""
dcMA.BoundText = ""
txtemp_code.Text = ""
FromDate.value = Date
ToDate.value = Date
FromDateH.value = ToHijriDate(Date)
ToDateH.value = ToHijriDate(Date)

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

 

 

Private Sub Dtp_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub



Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtID, "15062020006"


End Sub

Private Sub Date_Change()
dateH.value = ToHijriDate(Me.Date.value)
End Sub


Private Sub DateH_LostFocus()
VBA.Calendar = vbCalGreg
Me.Date.value = ToGregorianDate(Me.dateH.value)
End Sub

Private Sub dcEmp_Change()
Dim str As String
str = " select emp_id , Emp_code   from tblemployee  where emp_id = " & val(dcEmp.BoundText)
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
   txtemp_code.Text = IIf(IsNull(Rs_Temp("emp_code").value), "", Rs_Temp("emp_code").value)
End If
End Sub


Private Sub DCEmP_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
        Unload FrmEmployeeSearch
        FrmEmployeeSearch.lbltype = 2525
        FrmEmployeeSearch.show
End If
End Sub

Private Sub dcMA_Click(Area As Integer)
    Dim StrSQL As String
       If dcMA.BoundText = "" Then Exit Sub
       Set Rs_Temp = New ADODB.Recordset
       StrSQL = " select *  from tblschoolefile where id =    " & dcMA.BoundText
       Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       If (Rs_Temp.RecordCount > 0) Then
            txtMinistry.Text = IIf(IsNull(Rs_Temp("ministerno").value), "", Rs_Temp("ministerno").value)
       Else
              txtMinistry.Text = ""
       End If
End Sub

Private Sub dcMA_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "SSA"
            FrmSearch_BasicData.show
    End If
End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos
    Dim str As String

    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    
    If SystemOptions.UserInterface = ArabicInterface Then
        str = "  select id , name from TblSchooleFile  "
    Else
         str = "  select id , name from TblSchooleFile  "
    End If
    fill_combo dcMA, str
   
    
    If SystemOptions.UserInterface = ArabicInterface Then
            str = "   select  emp_id , emp_name  from tblEmployee  "
    Else
            str = "   select  emp_id , emp_nameE  from tblEmployee  "
    End If
       str = str & "  where  (BranchId=0 or BranchId is null or         BranchId in(" & Current_branchSql & "))"

    fill_combo dcEmp, str
   
    str = " select ID ,Name  from TblDurations  "
   fill_combo dcDuration, str
   
     If SystemOptions.UserInterface = EnglishInterface Then
            SetInterface Me
            ChangeLang
     End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " ⁄„·Ì«   Œ’Ì’ „‘—›Ì‰ ··„œ«—”  "
    LogTexte = " Open Window " & " Confirm  Violation "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
       
    
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblSupervisorAllocation "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    
    XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    FromDate.value = Date
    ToDate.value = Date
    FromDateH.value = ToHijriDate(FromDate.value)
    ToDateH.value = ToHijriDate(ToDate.value)
    Me.Date.value = Date
    Me.dateH.value = ToHijriDate(Date)
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

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«   Œ’Ì’ „‘—›Ì‰ ··„œ«—”  "
    LogTexte = " Exit Window " & "  Boxes Data "
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


Private Sub FromDate_Change()
FromDateH.value = ToHijriDate(FromDate.value)
End Sub


Private Sub Fromdateh_LostFocus()
VBA.Calendar = vbCalGreg
FromDate.value = ToGregorianDate(FromDateH.value)
End Sub




Private Sub ToDate_Change()
ToDateH.value = ToHijriDate(ToDate.value)
End Sub

Private Sub ToDateH_LostFocus()
VBA.Calendar = vbCalGreg
ToDate.value = ToGregorianDate(ToDateH.value)
End Sub

Private Sub txtemp_code_Change()

If txtemp_code.Text = "" Then Exit Sub
Dim str As String
str = " select emp_id , Emp_code   from tblemployee  where emp_code = " & txtemp_code.Text
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs_Temp.RecordCount > 0 Then
   dcEmp.BoundText = IIf(IsNull(Rs_Temp("emp_id").value), "", Rs_Temp("emp_id").value)
Else
   dcEmp.BoundText = ""
End If

End Sub

Private Sub txtemp_code_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmEmployeeSearch
        FrmEmployeeSearch.lbltype = 2525
        FrmEmployeeSearch.show
End If
End Sub

Private Sub txtMinistry_Change()
Dim val1, val2
If txtMinistry.Text = "" Then Exit Sub
Dim str As String
    str = " select *  from tblschoolefile where ministerNo =   '" & txtMinistry.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        dcMA.BoundText = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
     Else
        dcMA.BoundText = ""
    End If
End Sub

Private Sub txtMinistry_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "SSA"
            FrmSearch_BasicData.show
    End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   Œ’Ì’ „‘—›Ì‰ ··„œ«—” "
            Else
                Me.Caption = "Violation Types"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            
         '  Me.txtID.locked = True
            'Me.txtName.locked = True
          '  Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            C1Elastic2.Enabled = False
            C1Elastic3.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ⁄„·Ì«   Œ’Ì’ „‘—›Ì‰ ··„œ«—”( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ⁄„·Ì«   Œ’Ì’ „‘—›Ì‰ ··„œ«—”( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
             
            C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ⁄„·Ì«   Œ’Ì’ „‘—›Ì‰ ··„œ«—” (  ⁄œÌ· )"
            Else
                Me.Caption = "Violation Types(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
             
            C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
    
    txtID.Text = rs("IDSA").value
     dcBranch.BoundText = IIf(IsNull(rs("branchid").value), "", rs("branchid").value)
'     dcDuration.BoundText = IIf(IsNull(rs("DurationID").value), "", rs("DurationID").value)
    Me.Date.value = IIf(IsNull(rs("date").value), Date, rs("date").value)
    Me.dateH.value = IIf(IsNull(rs("date").value), ToHijriDate(Date), rs("date").value)
    
    
    Dim str As String
     str = " Select * from TblSupervisorAllocation_Details  where idsa =  " & val(txtID.Text)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
    Dim i As Integer
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst
        With Grid
            .Rows = Rs_Temp.RecordCount + 1
            For i = 1 To .Rows - 1
                  .TextMatrix(i, .ColIndex("Serial")) = i
                 .TextMatrix(i, .ColIndex("schoolFileID")) = IIf(IsNull(Rs_Temp("schoolFileID").value), "", (Rs_Temp("schoolFileID").value))
                 .TextMatrix(i, .ColIndex("SchoolFileName")) = IIf(IsNull(Rs_Temp("SchoolFileName").value), "", (Rs_Temp("SchoolFileName").value))
                 .TextMatrix(i, .ColIndex("emp_id")) = IIf(IsNull(Rs_Temp("emp_id").value), "", (Rs_Temp("emp_id").value))
                 .TextMatrix(i, .ColIndex("emp_name")) = IIf(IsNull(Rs_Temp("emp_name").value), "", (Rs_Temp("emp_name").value))
                 .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(Rs_Temp("FromDate").value), Date, (Rs_Temp("FromDate").value))
                 .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(Rs_Temp("FromDateH").value), ToHijriDate(Date), (Rs_Temp("FromDateH").value))
                 .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(Rs_Temp("ToDate").value), Date, (Rs_Temp("ToDate").value))
                 .TextMatrix(i, .ColIndex("TODateH")) = IIf(IsNull(Rs_Temp("TODateH").value), ToHijriDate(Date), (Rs_Temp("TODateH").value))
                 .TextMatrix(i, .ColIndex("ministryno")) = IIf(IsNull(Rs_Temp("ministryno").value), "", (Rs_Temp("ministryno").value))
                 Rs_Temp.MoveNext
            Next
        End With
    End If

    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub




Private Sub TxtName_GotFocus()
On Error Resume Next
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
 SwitchKeyboardLang LANG_ENGLISH
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

Function CuurentLogdata(Optional Currentmode As String)
   
  

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap


    If dcBranch.BoundText = "" Then
        MsgBox (" «Œ — «·›—⁄ «Ê·«  ")
        Exit Sub
    End If

     If dcDuration.BoundText = "" Then
        MsgBox (" «Œ — «·”‰… «·œ—«”Ì… «Ê·«  ")
        Exit Sub
    End If

    If Grid.Rows = 1 Then
            MsgBox (" ﬁ„ »⁄„·ÌÂ  Œ’Ì’ „‘—›Ì‰ «Ê·«  ")
        Exit Sub
    End If

    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
            Case "N"
            rs.AddNew
                    txtID.Text = CStr(new_id("TblSupervisorAllocation", "IDSA", "", True))
            Case "E"
              '  StrSQL = "select * From  TblViolationTypes where Name='" & Trim(txtName.text) & "'"
               
        End Select

        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg.Text
            Case "N"
               rs("IDSA").value = val(txtID.Text)
               rs("UserID").value = user_id
               rs("CreationDate").value = Date
            Case "E"
              StrSQL = "delete From TblSupervisorAllocation_Details where  IDSA =" & val(txtID.Text)
              Cn.Execute StrSQL, , adExecuteNoRecords
        End Select
          
          rs("branchid") = dcBranch.BoundText
          rs("date") = Me.Date.value
          rs("DateH") = Me.dateH.value
          rs("DurationID") = IIf(dcDuration.BoundText = "", Null, dcDuration.BoundText)
          rs.update
        
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open "TblSupervisorAllocation_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Dim i As Integer
        
        With Grid
            For i = 1 To Grid.Rows - 1
               Rs_Temp.AddNew
               Rs_Temp("id").value = CStr(new_id("TblSupervisorAllocation_Details", "id", "", True))
               Rs_Temp("idsa").value = val(txtID.Text)
               Rs_Temp("schoolFileID").value = .TextMatrix(i, .ColIndex("schoolFileID"))
               Rs_Temp("schoolFileName").value = .TextMatrix(i, .ColIndex("schoolFileName"))
               Rs_Temp("emp_id").value = .TextMatrix(i, .ColIndex("emp_id"))
               Rs_Temp("emp_name").value = .TextMatrix(i, .ColIndex("emp_name"))
               Rs_Temp("FromDate").value = .TextMatrix(i, .ColIndex("FromDate"))
               Rs_Temp("FromDateH").value = .TextMatrix(i, .ColIndex("FromDateH"))
               Rs_Temp("ToDate").value = .TextMatrix(i, .ColIndex("ToDate"))
               Rs_Temp("TODateH").value = .TextMatrix(i, .ColIndex("TODateH"))
               Rs_Temp("ministryno").value = .TextMatrix(i, .ColIndex("ministryno"))
               Rs_Temp.update
               
                Dim str3 As String
                str3 = "  select * from TblSchooleFile  where id  =  " & val(.TextMatrix(i, .ColIndex("schoolFileID")))
                Set Rs_Temp2 = New ADODB.Recordset
                Rs_Temp2.Open str3, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs_Temp2.RecordCount > 0 Then
                        Rs_Temp2("Emp_ID").value = IIf(IsNull(.TextMatrix(i, .ColIndex("emp_id"))), Null, .TextMatrix(i, .ColIndex("emp_id")))
                        Rs_Temp2.update
                End If
               
             Next
        End With
      
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & CHR(13)
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
            rs.find "ID='" & val(txtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtID.Text <> "" Then

    
        Msg = "”Ì „ Õ–› «·»Ì«‰«  —ﬁ„ " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblSupervisorAllocation  where  IDSA =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                   
                StrSQL = "delete From TblSupervisorAllocation_Details  where  IDSA =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                    
                   
                   StrSQL = "SELECT  *  From TblSupervisorAllocation  "
                   Set rs = New ADODB.Recordset
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «· ⁄ÿ· "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
  '  Dim Wrap As String
  '  On Error GoTo ErrTrap
  '  Set TTP = New clstooltip
  '  Wrap = Chr(13) + Chr(10)
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  «À»«  «· ⁄ÿ·  ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ·  ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «À»«  «· ⁄ÿ· " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ·  ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «À»«  «· ⁄ÿ·  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «À»«  «· ⁄ÿ· " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰« «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ «À»«  «· ⁄ÿ· " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
''    With TTP
''        .Create Me.hWnd, "»Ì«‰«  «· ⁄ÿ·", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«   «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
'    End With
'
'    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub




