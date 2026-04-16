VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmUnitÒReceive 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÃ“ ÊÕœ« "
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12540
   Icon            =   "FrmUnitRecerive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   12540
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   3000
      Width           =   12495
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   11
         Left            =   120
         TabIndex        =   73
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
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
         ButtonImage     =   "FrmUnitRecerive.frx":038A
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Height          =   315
         Left            =   10200
         TabIndex        =   74
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ÊÕœ…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   11
         Left            =   9240
         TabIndex        =   78
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œÊ—"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   13
         Left            =   6120
         TabIndex        =   77
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·√Ã— «·ÌÊ„Ì"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   16
         Left            =   2520
         TabIndex        =   76
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÊÕœ…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   22
         Left            =   10800
         TabIndex        =   75
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2280
      Width           =   12495
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœ«"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘»ﬂ…"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   3195
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
         Width           =   1995
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   51
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
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
         ButtonImage     =   "FrmUnitRecerive.frx":0924
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   58
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œ›⁄… «·„ﬁœ„…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   12
         Left            =   10800
         TabIndex        =   57
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2E9E9&
      Height          =   1215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1080
      Width           =   12495
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   795
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   375
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   915
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   375
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   375
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   13
         Left            =   120
         TabIndex        =   41
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
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
         ButtonImage     =   "FrmUnitRecerive.frx":0EBE
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ«  ⁄‰ «·⁄ﬁœ"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   9
         Left            =   4080
         TabIndex        =   49
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«Ì«„"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   10
         Left            =   4680
         TabIndex        =   48
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Â«Ì… «·«ÌÃ«—"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   14
         Left            =   11400
         TabIndex        =   47
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÌÊ„"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   18
         Left            =   8880
         TabIndex        =   46
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·ﬁœÊ„"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   19
         Left            =   11400
         TabIndex        =   45
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”«⁄…"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   17
         Left            =   10920
         TabIndex        =   44
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”«⁄…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   21
         Left            =   10920
         TabIndex        =   43
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÌÊ„"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   15
         Left            =   8880
         TabIndex        =   42
         Top             =   720
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   1095
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   12495
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   3855
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   10
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
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
         ButtonImage     =   "FrmUnitRecerive.frx":1458
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«À»« "
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   33
         Left            =   10920
         TabIndex        =   26
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   34
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«—ﬁ«„ «·« ’«· .Ã"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   36
         Left            =   10920
         TabIndex        =   24
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„” √Ã—"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   38
         Left            =   8160
         TabIndex        =   23
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " .„"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   40
         Left            =   7440
         TabIndex        =   22
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " .⁄"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   41
         Left            =   3720
         TabIndex        =   21
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   16200
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   15720
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   16320
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8580
      TabIndex        =   0
      Top             =   3960
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   16200
      TabIndex        =   1
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   15840
      TabIndex        =   9
      Top             =   1920
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   480
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4320
      Width           =   8745
      _cx             =   15425
      _cy             =   953
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
      Appearance      =   0
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7230
         TabIndex        =   60
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   6375
         TabIndex        =   61
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   5535
         TabIndex        =   62
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   4680
         TabIndex        =   63
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   3825
         TabIndex        =   64
         Top             =   75
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
         Left            =   0
         TabIndex        =   65
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   855
         TabIndex        =   66
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   67
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   68
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â"
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
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   15090
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11325
      TabIndex        =   7
      Top             =   4035
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   6
      Top             =   3870
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   5
      Top             =   3870
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   3900
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   3900
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   16350
      TabIndex        =   2
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmUnitÒReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Resize_Form Me
End Sub
