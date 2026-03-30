VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form all_alarms 
   Caption         =   "     ‰»ÌÂ«  ‘∆Ê‰ «·„ÊŸ›Ì‰     "
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "all_alarms.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   10650
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8280
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   10650
      _cx             =   18785
      _cy             =   14605
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
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
      Begin ALLButtonS.ALLButton x1 
         Height          =   570
         Left            =   0
         TabIndex        =   4
         Top             =   1875
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1005
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   585
         Left            =   0
         TabIndex        =   5
         Top             =   3510
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton x2 
         Height          =   765
         Left            =   0
         TabIndex        =   6
         Top             =   2580
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1349
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton4 
         Height          =   675
         Left            =   0
         TabIndex        =   7
         Top             =   4350
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton6 
         Height          =   570
         Left            =   0
         TabIndex        =   8
         Top             =   5280
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1005
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton7 
         Height          =   600
         Left            =   0
         TabIndex        =   9
         Top             =   5910
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1058
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton8 
         Height          =   810
         Left            =   0
         TabIndex        =   23
         Top             =   6600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1429
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   810
         Left            =   0
         TabIndex        =   24
         Top             =   7320
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1429
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":00D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   570
         Left            =   0
         TabIndex        =   26
         Top             =   1320
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1005
         BTYPE           =   3
         TX              =   "⁄—÷"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "all_alarms.frx":00EC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2730
         Index           =   13
         Left            =   300
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   450
         Width           =   10380
         _cx             =   18309
         _cy             =   4815
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
         Begin VB.TextBox TxtItemQty 
            Alignment       =   1  'Right Justify
            Height          =   0
            Index           =   2
            Left            =   15
            MaxLength       =   5
            TabIndex        =   34
            Top             =   1170
            Width           =   0
         End
         Begin VB.TextBox TxtItemPrice 
            Alignment       =   1  'Right Justify
            Height          =   0
            Index           =   2
            Left            =   15
            MaxLength       =   5
            TabIndex        =   33
            Top             =   1170
            Visible         =   0   'False
            Width           =   0
         End
         Begin VB.TextBox TXTLCNO 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6945
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   210
            Width           =   1755
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   5295
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1020
            Width           =   3420
         End
         Begin VB.TextBox TXTValue 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   7005
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1530
            Width           =   1755
         End
         Begin VB.TextBox txtNameE 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1020
            Width           =   3090
         End
         Begin MSDataListLib.DataCombo DCPreFix 
            Height          =   315
            Left            =   5295
            TabIndex        =   35
            Top             =   225
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   979
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCLC 
            Height          =   315
            Left            =   210
            TabIndex        =   36
            Top             =   210
            Width           =   3090
            _ExtentX        =   5450
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
         Begin MSDataListLib.DataCombo DCBank 
            Height          =   315
            Left            =   5295
            TabIndex        =   37
            Top             =   660
            Width           =   3420
            _ExtentX        =   6033
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
         Begin MSDataListLib.DataCombo DcBranch 
            Height          =   315
            Left            =   210
            TabIndex        =   38
            Top             =   630
            Width           =   3090
            _ExtentX        =   5450
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
            Caption         =   "«·”⁄—"
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
            Height          =   990
            Index           =   49
            Left            =   15
            TabIndex        =   48
            Top             =   75
            Visible         =   0   'False
            Width           =   30
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
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
            Height          =   990
            Index           =   50
            Left            =   15
            TabIndex        =   47
            Top             =   75
            Width           =   30
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊÕœÂ"
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
            Height          =   1050
            Index           =   51
            Left            =   15
            TabIndex        =   46
            Top             =   120
            Width           =   30
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   450
            Index           =   35
            Left            =   8565
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   300
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄"
            Height          =   225
            Index           =   4
            Left            =   3735
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   300
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   270
            Index           =   25
            Left            =   8595
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1125
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»‰ﬂ"
            Height          =   255
            Index           =   6
            Left            =   9225
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   765
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   225
            Index           =   9
            Left            =   8745
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1680
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   330
            Index           =   29
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   675
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ì"
            Height          =   330
            Index           =   28
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1050
            Width           =   1635
         End
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰»ÌÂ«  «⁄Ì«œ „Ì·«œ «·⁄„·«¡"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1320
         Width           =   5370
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰»ÌÂ«    ··⁄ﬁÊœ «· Ì ” ‰ ÂÌ Œ·«· › —Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3975
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   7485
         Width           =   5385
      End
      Begin VB.Image Image1 
         Height          =   1365
         Left            =   150
         Picture         =   "all_alarms.frx":0108
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰»ÌÂ«  › —… «·«Œ »«— ··⁄ﬁÊœ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4035
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   6765
         Width           =   5370
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "      ‰»ÌÂ«  ‘∆Ê‰ «·„ÊŸ›Ì‰     "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1365
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   0
         Width           =   10530
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì  ⁄œœ «·«ﬁ«„«  «· Ì ” ‰ ÂÌ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1875
         Width           =   5370
      End
      Begin VB.Label d1 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1890
         Width           =   1785
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì  ⁄œœ «·ÃÊ«“«   «· Ì ” ‰ ÂÌ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   3645
         Width           =   5370
      End
      Begin VB.Label d2 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3645
         Width           =   1830
      End
      Begin VB.Label d3 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì  ⁄œœ  —Œ’ «·⁄„· «· Ì ” ‰ ÂÌ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2700
         Width           =   5370
      End
      Begin VB.Label d4 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   4470
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì  ⁄œœ «·ÂÊÌ«  «· Ì ” ‰ ÂÌ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   4350
         Width           =   5370
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ã„«·Ì  ⁄œœ —Œ’ «·ﬁÌ«œ… «·„‰ ÂÌ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   5265
         Width           =   5370
      End
      Begin VB.Label d6 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   5370
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰»ÌÂ«  «·«Ã«“« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   5925
         Width           =   5370
      End
   End
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "⁄—÷"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "all_alarms.frx":0F84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì  ⁄œœ  «·„ÊŸ›Ì‰  «· Ì ” ‰ ÂÌ  √„Ì‰« Â„"
      Height          =   855
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label d5 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   615
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "all_alarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Askinterval As String
Dim Askcount As Integer

Private Sub ALLButton1_Click()
FrmContractExam.mIndex = 1
FrmContractExam.show

End Sub

Private Sub ALLButton2_Click()
    FrmEmpExpir1.show
End Sub

Private Sub ALLButton3_Click()
FrmEmpExpir4.mIndex = 1
FrmEmpExpir4.show
End Sub

Private Sub ALLButton4_Click()
    FrmEmpExpir4.show
End Sub

Private Sub ChangeLang()
    Me.Caption = "            Today Alarms           "
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Total No of Expired Residence"
    Label4.Caption = "Total No of Expired Passport"
    Label5.Caption = "Total No of Expired Work License"
    Label7.Caption = "Total No of Expired ID"
    Label9.Caption = "Total No of Expired Insurance"
    Label8.Caption = "Expired Contract Test"
 Label6.Caption = "Vacation Alarms"
        
      Label3.Caption = "Total No of Expired Driver License"
      
   x1.Caption = "View"
    ALLButton2.Caption = "View"
   X2.Caption = "View"
    ALLButton4.Caption = "View"
    ALLButton5.Caption = "View"
    
        ALLButton6.Caption = "View"
          ALLButton7.Caption = "View"
          ALLButton8.Caption = "View"
          
End Sub

Private Sub ALLButton6_Click()
    FrmEmpExpir5.show
End Sub

Private Sub ALLButton7_Click()
FrmAlarmVacation.show

End Sub

Private Sub ALLButton8_Click()
FrmContractExam.mIndex = 0
FrmContractExam.show

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    ' My_SQL = "select * From TblEmployee  where DateEndPasp < getdate()"
    ' My_SQL = "SELECT     * from dbo.TblEmployee Where (Month(DateEndPasp) <= Month(GetDate())) And (year(DateEndPasp) <= year(GetDate()))"
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_Expirepas", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_Expirepas", 0)
 
'  My_SQL = "SELECT     * from dbo.TblEmployee Where (  NOT (dbo.TblEmployee.NumPasp IS NULL ) ) and  DateEndPasp<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
My_SQL = "SELECT     * from dbo.TblEmployee Where (  NOT (dbo.TblEmployee.NumPasp IS NULL ) ) and  DateEndPasp<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    d2.Caption = rs.RecordCount
    rs.Close

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireEkama", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireEkama", 0)
'    My_SQL = "SELECT     * from dbo.emp_all_details Where   (NOT (dbo.emp_all_details.NumEkama IS NULL))  AND DateEndekama<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
My_SQL = "SELECT     * from dbo.emp_all_details Where  dbo.emp_all_details.NationlID <>1 and  (NOT (dbo.emp_all_details.NumEkama IS NULL))  AND DateEndekama<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"

   
  
    My_SQL = My_SQL & " order by DateEndekama,fullcode"
 
 
   
   
   ' My_SQL = My_SQL & " order by DateEndekama,fullcode"
  
  
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    d1.Caption = rs.RecordCount
    rs.Close

    'My_SQL = "SELECT     * from dbo.TblEmployee Where (Month(DateEndLinc) <= Month(GetDate())) And (year(DateEndLinc) <= year(GetDate()))"
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireLicence", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireLicence", 0)
    My_SQL = "SELECT     * from dbo.TblEmployee Where  (NOT (NumLicn IS NULL)) and DateEndLinc<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"

    ' My_SQL = "select * From TblEmployee  where DateEndLinc < getdate()"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    d3.Caption = rs.RecordCount

    rs.Close
    'My_SQL = "select * From TblEmployee  where dateendpoket < getdate()"
    '  My_SQL = "SELECT     * from dbo.TblEmployee Where (Month(dateendpoket) <= Month(GetDate())) And (year(dateendpoket) <= year(GetDate()))"
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_Expirepoket", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_Expirepoket", 0)
    My_SQL = "SELECT     * from dbo.TblEmployee Where dateendpoket<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
My_SQL = "SELECT     * from dbo.TblEmployee Where   (NOT (dbo.TblEmployee.NumPoket IS NULL )  )  and dateendpoket<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
My_SQL = "SELECT     * from dbo.TblEmployee Where   (NOT (dbo.TblEmployee.NumPoket IS NULL ))  and dateendpoket<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    d4.Caption = rs.RecordCount

    rs.Close
    '    My_SQL = "select * From TblEmployee  where dateendpoket < getdate()"
    My_SQL = "SELECT     * from dbo.TblEmployee Where (Month(dateendpoket) <= Month(GetDate())) And (year(dateendpoket) <= year(GetDate()))"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    d5.Caption = rs.RecordCount
    
        rs.Close
   '
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireLicence", "D")
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireLicence", 0)
    My_SQL = "SELECT     * from dbo.TblEmployee Where DriverLicenseend<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    My_SQL = "SELECT     * from dbo.TblEmployee Where   (NOT (dbo.TblEmployee.DriverLicense IS NULL )) and DriverLicenseend<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    d6.Caption = rs.RecordCount

    rs.Close
    
    
    
    

End Sub

Private Sub x1_Click()
    FrmEmpExpir2.show
End Sub

Private Sub x2_Click()
    FrmEmpExpir3.show
End Sub
