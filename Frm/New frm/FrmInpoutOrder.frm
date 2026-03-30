VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInpoutOrder 
   Caption         =   "«–‰ «÷«›… "
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   HelpContextID   =   100
   Icon            =   "FrmInpoutOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmInpoutOrder.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   12225
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   6945
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   12225
      _cx             =   21564
      _cy             =   12250
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
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
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
      GridRows        =   5
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmInpoutOrder.frx":2B2C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ImpulseButton.ISButton CmdInfo 
         Height          =   600
         Left            =   11610
         TabIndex        =   46
         Top             =   15
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmInpoutOrder.frx":2B99
         ButtonImageHover=   "FrmInpoutOrder.frx":3873
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5940
         Width           =   12195
         _cx             =   21511
         _cy             =   767
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
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   360
            Left            =   6255
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   330
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3345
            TabIndex        =   36
            Top             =   75
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   255
            Index           =   24
            Left            =   7275
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   90
            Width           =   690
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   8025
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   30
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   50
            Left            =   9285
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   90
            Width           =   750
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10065
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   30
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   255
            Index           =   23
            Left            =   1125
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   90
            Width           =   195
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5925
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   30
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   255
            Index           =   1
            Left            =   5070
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   90
            Width           =   840
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   270
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   90
            Width           =   1050
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   270
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   90
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   270
            Index           =   0
            Left            =   2265
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   90
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   3
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   105
            Width           =   780
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   6390
         Width           =   12195
         _cx             =   21511
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
            Height          =   375
            Index           =   0
            Left            =   10890
            TabIndex        =   9
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9540
            TabIndex        =   10
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   8145
            TabIndex        =   11
            Top             =   90
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   6810
            TabIndex        =   12
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   5445
            TabIndex        =   13
            Top             =   90
            Width           =   1260
            _ExtentX        =   2223
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   4080
            TabIndex        =   14
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   30
            TabIndex        =   17
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   2670
            TabIndex        =   15
            Top             =   90
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   375
            Left            =   1350
            TabIndex        =   16
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1695
         Index           =   5
         Left            =   15
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   630
         Width           =   12195
         _cx             =   21511
         _cy             =   2990
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
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   8400
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   60
            Width           =   2025
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   600
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   960
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   120
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox Txt_EXport 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   2370
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   1200
            Width           =   930
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   120
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   840
            Width           =   750
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   1215
            Width           =   750
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6870
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   60
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   60
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   510
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   510
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   510
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   6345
            TabIndex        =   3
            Top             =   840
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   6345
            TabIndex        =   5
            Top             =   1200
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   7890
            TabIndex        =   1
            Top             =   480
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            Format          =   82444289
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   420
            Left            =   5520
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   795
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmInpoutOrder.frx":454D
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   390
            Left            =   420
            TabIndex        =   152
            Top             =   1200
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕÊÌ· ≈·Ì ›« Ê—…"
            BackColor       =   12632256
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„’—Ê›«  «·«Œ—Ï"
            Height          =   255
            Left            =   3405
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   1320
            Width           =   2355
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   210
            Index           =   4
            Left            =   10485
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1215
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   345
            Index           =   5
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   135
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ"
            Height          =   285
            Index           =   6
            Left            =   10485
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   855
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   300
            Index           =   8
            Left            =   10485
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   105
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ «·ÂÃ—Ì"
            Height          =   270
            Index           =   9
            Left            =   6330
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   480
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·›« Ê—…"
            Height          =   270
            Index           =   7
            Left            =   10485
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   300
            Index           =   10
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   855
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·Œ’„"
            Height          =   285
            Index           =   11
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   480
            Visible         =   0   'False
            Width           =   1305
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   600
         Index           =   6
         Left            =   15
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   15
         Width           =   11580
         _cx             =   20426
         _cy             =   1058
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
         Caption         =   "«–‰ «÷«›… «’‰«› „’‰⁄Â ··«‰ «Ã «· «„"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   345
            Left            =   5595
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   90
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   315
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   105
            Visible         =   0   'False
            Width           =   405
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   390
            Left            =   3705
            TabIndex        =   63
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpoutOrder.frx":48E7
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1965
            TabIndex        =   21
            Top             =   105
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpoutOrder.frx":4C81
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
            Left            =   1050
            TabIndex        =   22
            Top             =   105
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpoutOrder.frx":501B
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
            Left            =   2805
            TabIndex        =   20
            Top             =   105
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpoutOrder.frx":53B5
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
            Left            =   135
            TabIndex        =   23
            Top             =   105
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpoutOrder.frx":574F
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   390
            Left            =   4650
            TabIndex        =   64
            Top             =   90
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmInpoutOrder.frx":5AE9
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   3585
         Left            =   15
         TabIndex        =   43
         Top             =   2340
         Width           =   12195
         _cx             =   21511
         _cy             =   6324
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
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…|≈” ﬁÿ«⁄«  ⁄·Ï «·›« Ê—…"
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
         DogEars         =   0   'False
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   1
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "FrmInpoutOrder.frx":6083
         Picture(1)      =   "FrmInpoutOrder.frx":641D
         Flags(1)        =   3
         Flags(2)        =   3
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3120
            Index           =   0
            Left            =   45
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   45
            Width           =   12105
            _cx             =   21352
            _cy             =   5503
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
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   1
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
            GridRows        =   3
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpoutOrder.frx":67B7
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   1980
               Left            =   30
               TabIndex        =   45
               Top             =   735
               Width           =   12045
               _cx             =   21246
               _cy             =   3492
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
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   13
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInpoutOrder.frx":6825
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin MSComctlLib.Toolbar TBar 
               Height          =   600
               Left            =   495
               TabIndex        =   47
               Top             =   2730
               Width           =   10920
               _ExtentX        =   19262
               _ExtentY        =   1058
               ButtonWidth     =   609
               ButtonHeight    =   953
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   4
               Left            =   30
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   30
               Width           =   12045
               _cx             =   21246
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
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   2
               ChildSpacing    =   1
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
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   53
                  Top             =   315
                  Width           =   1800
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   1665
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   315
                  Width           =   1380
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   3045
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   315
                  Width           =   2730
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   585
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   315
                  Width           =   1080
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   7680
                  TabIndex        =   54
                  Top             =   315
                  Width           =   2790
                  _ExtentX        =   4921
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   10515
                  TabIndex        =   55
                  Top             =   315
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   345
                  Left            =   45
                  TabIndex        =   56
                  Top             =   315
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
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
                  ButtonImage     =   "FrmInpoutOrder.frx":6A11
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   270
                  Index           =   31
                  Left            =   10515
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   30
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   270
                  Index           =   30
                  Left            =   7680
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   30
                  Width           =   2790
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   270
                  Index           =   29
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   30
                  Width           =   1800
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   270
                  Index           =   28
                  Left            =   3045
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   30
                  Width           =   2730
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   270
                  Index           =   27
                  Left            =   1665
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   30
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   270
                  Index           =   26
                  Left            =   585
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   30
                  Width           =   1080
               End
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   360
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   2730
               Width           =   450
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3120
            Index           =   2
            Left            =   12840
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   45
            Width           =   12105
            _cx             =   21352
            _cy             =   5503
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
            BackColor       =   255
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   0
            ChildSpacing    =   0
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
            GridRows        =   3
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpoutOrder.frx":6DAB
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1245
               Index           =   10
               Left            =   0
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   1875
               Width           =   12105
               _cx             =   21352
               _cy             =   2196
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
               AutoSizeChildren=   8
               BorderWidth     =   1
               ChildSpacing    =   1
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
               GridRows        =   4
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmInpoutOrder.frx":6E17
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   375
                  Index           =   14
                  Left            =   15
                  TabIndex        =   67
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   9735
                  _cx             =   17171
                  _cy             =   661
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   285
                     Index           =   2
                     Left            =   8805
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   60
                     Width           =   885
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   285
                     Left            =   2610
                     TabIndex        =   69
                     Top             =   60
                     Width           =   1560
                     _ExtentX        =   2752
                     _ExtentY        =   503
                     ButtonStyle     =   1
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
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
                     Height          =   285
                     Index           =   18
                     Left            =   4185
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   60
                     Width           =   870
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
                     Height          =   285
                     Index           =   16
                     Left            =   5070
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   60
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Ìﬂ« "
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
                     Height          =   285
                     Index           =   17
                     Left            =   7650
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   60
                     Width           =   1080
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   19
                     Left            =   6750
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   60
                     Width           =   870
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   1080
                  Left            =   1740
                  TabIndex        =   74
                  Top             =   405
                  Width           =   8010
                  _cx             =   14129
                  _cy             =   1905
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
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmInpoutOrder.frx":6E91
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1335
               Index           =   7
               Left            =   0
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   540
               Width           =   12105
               _cx             =   21352
               _cy             =   2355
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
               AutoSizeChildren=   8
               BorderWidth     =   1
               ChildSpacing    =   1
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
               GridRows        =   3
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmInpoutOrder.frx":6FC5
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   765
                  Left            =   1815
                  TabIndex        =   76
                  Top             =   465
                  Width           =   7935
                  _cx             =   13996
                  _cy             =   1349
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
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmInpoutOrder.frx":7035
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   345
                  Index           =   13
                  Left            =   15
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   1245
                  Width           =   9735
                  _cx             =   17171
                  _cy             =   609
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
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   60
                     Width           =   630
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «· ﬁ”Ìÿ"
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
                     Height          =   225
                     Index           =   42
                     Left            =   2250
                     RightToLeft     =   -1  'True
                     TabIndex        =   91
                     Top             =   60
                     Width           =   960
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Top             =   60
                     Width           =   765
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê· ﬁ”ÿ"
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
                     Height          =   225
                     Index           =   40
                     Left            =   4020
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Top             =   60
                     Width           =   690
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   4740
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   60
                     Width           =   330
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√ﬁ”«ÿ"
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
                     Height          =   225
                     Index           =   38
                     Left            =   5100
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   60
                     Width           =   915
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   6030
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   60
                     Width           =   585
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»·€ «·ﬂ·Ï"
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
                     Height          =   225
                     Index           =   36
                     Left            =   6645
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   60
                     Width           =   870
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   8280
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   60
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·›«∆œ…"
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
                     Height          =   225
                     Index           =   34
                     Left            =   8940
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   60
                     Width           =   750
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·›«∆œ…"
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
                     Height          =   225
                     Index           =   35
                     Left            =   7830
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   60
                     Width           =   435
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   7545
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   60
                     Width           =   270
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   1980
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   60
                     Width           =   240
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   60
                     Width           =   225
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„»œ∆Ì…"
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
                     Height          =   225
                     Index           =   37
                     Left            =   270
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   60
                     Width           =   1005
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   435
                  Index           =   12
                  Left            =   15
                  TabIndex        =   93
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   9735
                  _cx             =   17171
                  _cy             =   767
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   330
                     Index           =   1
                     Left            =   8820
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   30
                     Width           =   870
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Index           =   1
                     Left            =   7185
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   30
                     Width           =   1110
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Index           =   1
                     Left            =   5280
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   30
                     Width           =   1275
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   360
                     Left            =   1815
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   15
                     Width           =   855
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   435
                     Left            =   45
                     TabIndex        =   98
                     Top             =   -15
                     Width           =   1740
                     _ExtentX        =   3069
                     _ExtentY        =   767
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
                     BackColor       =   14871017
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmInpoutOrder.frx":7106
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin MSComCtl2.DTPicker DtpDelayDate 
                     Height          =   345
                     Left            =   2835
                     TabIndex        =   99
                     Top             =   30
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   609
                     _Version        =   393216
                     Format          =   82444289
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   330
                     Index           =   14
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   75
                     Width           =   540
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   345
                     Index           =   15
                     Left            =   8310
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   75
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   390
                     Index           =   21
                     Left            =   4425
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   -15
                     Width           =   765
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   540
               Index           =   11
               Left            =   0
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   0
               Width           =   12105
               _cx             =   21352
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
               Begin MSDataListLib.DataCombo DcboCurrency 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   112
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   8940
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   90
                  Width           =   720
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   5280
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   60
                  Width           =   1275
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   7215
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   60
                  Width           =   1095
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   2790
                  TabIndex        =   107
                  Top             =   105
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   345
                  Index           =   2
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   6555
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   90
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   8295
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   90
                  Width           =   450
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„·…"
                  Height          =   225
                  Index           =   20
                  Left            =   2190
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   525
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3120
            Index           =   15
            Left            =   13140
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   45
            Width           =   12105
            _cx             =   21352
            _cy             =   5503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
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
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   1
            ChildSpacing    =   1
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
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   7
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpoutOrder.frx":74A0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   630
               Left            =   15
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   129
               Top             =   2475
               Width           =   12075
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   18
               Left            =   15
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   1545
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   225
                  Left            =   8505
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6135
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   225
                  Index           =   49
                  Left            =   4875
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   315
                  Index           =   43
                  Left            =   7230
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   510
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   5775
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   360
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   17
               Left            =   15
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   1035
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   255
                  Left            =   8505
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6135
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   225
                  Index           =   33
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   41
                  Left            =   7230
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "$"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   5820
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   16
               Left            =   15
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   525
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   435
                  Left            =   2580
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   30
                  Width           =   585
               End
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2025
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   60
                  Width           =   360
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   32
                  Left            =   1605
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   90
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   39
                  Left            =   2415
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   90
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   46
                  Left            =   1905
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   90
                  Width           =   120
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   8
               Left            =   15
               TabIndex        =   130
               TabStop         =   0   'False
               Top             =   15
               Width           =   12075
               _cx             =   21299
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
               Appearance      =   5
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
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
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2025
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   75
                  Width           =   360
               End
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   225
                  Left            =   2685
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   120
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   225
                  Index           =   25
                  Left            =   1605
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   120
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   225
                  Index           =   22
                  Left            =   2415
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   135
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   45
                  Left            =   1905
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   120
                  Width           =   120
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… √Ì… „·«ÕŸ«  ⁄·Ï «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   405
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   2055
               Width           =   12075
            End
         End
      End
   End
End
Attribute VB_Name = "FrmInpoutOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim NewGrid As ClsGrid
Dim TTP As clstooltip
Dim BuyReport As ClsBuyReport
Dim cSearchDcbo(3) As clsDCboSearch

Public BolPrint As Boolean
Dim WithEvents m_MnuShowNewItemsPrices As Menu
Attribute m_MnuShowNewItemsPrices.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuShowItemCostEffect As Menu
Attribute m_MenuShowItemCostEffect.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim general_noteid As Long
Dim RsNotesGeneral As ADODB.Recordset

Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If

End Sub

Private Sub ChkTaxAdd_Click()

    If ChkTaxAdd.value = Checked Then
        TxtTaxAddValue.Enabled = True
        lbl(39).Enabled = True
        lbl(46).Enabled = True
    Else
        TxtTaxAddValue.text = ""
        TxtTaxAddValue.Enabled = False
        lbl(39).Enabled = False
        lbl(46).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxSerivce_Click()
    On Error GoTo ErrTrap

    If ChkTaxSerivce.value = Checked Then
        TxtTaxServiceValue.Enabled = True
        lbl(43).Enabled = True
        lbl(47).Enabled = True
    Else
        TxtTaxServiceValue.text = ""
        TxtTaxServiceValue.Enabled = False
        lbl(43).Enabled = False
        lbl(47).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxStamp_Click()

    If ChkTaxStamp.value = Checked Then
        TxtTaxStampValue.Enabled = True
        lbl(41).Enabled = True
        lbl(48).Enabled = True
    Else
        TxtTaxStampValue.text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String

    BolPrint = True

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type= 20"))
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
            DCboStoreName.BoundText = intDef
            XPTab301.CurrTab = 0
            Fg.SetFocus
            Fg.Col = Fg.ColIndex("Code")
            Fg.Row = Fg.Rows - 1

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                If AvailableDeal = False Then
                    Exit Sub
                End If
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            Me.DcboBox.BoundText = 1

        Case 2

            If Me.TxtModFlg.text = "N" Then
             
                If TxtNoteSerial.text = "" Then
                    If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                    Else
                       
                        If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                        Else
                            TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
                        End If
                    End If
                End If
        
                If TxtNoteSerial1.text = "" Then
                    If Voucher_coding(val(my_branch), XPDtbBill.value, 19, 250) = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ œ›⁄ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                       
                        If Voucher_coding(val(my_branch), XPDtbBill.value, 19, 250) = "" Then
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 19, 250)
                        End If
                    End If
                End If
             
            End If
             
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            If Text1.text <> "" Then
                Msg = " „  ÕÊÌ· Â–« «·«–‰ »—ﬁ„ ›« Ê—… ›·« Ì„ﬂ‰ﬂ «·Õ–› " & Space$(5) & Text1.text
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = PurchaseTransaction
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡"
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModeless, mdifrmmain
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’À… »‘«‘… ›« Ê—… «·‘—«¡ «·Õ«·Ì…"
                Msg = Msg & Chr(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                m_FrmSearch.Visible = True
                m_FrmSearch.ZOrder 0
                m_FrmSearch.SetFocus
            End If

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
            End If

            If BolPrint = False Then
                Exit Sub
            End If

            printing
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdCheque_Click()

    If Me.TxtModFlg.text = "R" Then
        Exit Sub
    End If

    Load FrmChecks
    FrmChecks.TxtModFlg.text = Me.TxtModFlg.text
    FrmChecks.XPTxtBillID.text = Me.XPTxtBillID.text
    Set FrmChecks.PutFg = Me.FgCheques
    FrmChecks.show vbModal
    SumChecks

End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .Rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .Rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .Rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String

    If Text1.text <> "" Then
        Msg = " „  ÕÊÌ· Â–« «·«–‰ »—ﬁ„ ›« Ê—…  " & Space$(5) & Text1.text
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    Set Frm = New FrmBillBuy

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .CboPaymentType.ListIndex = 0 ' CboPaymentType.ListIndex
        .Text1.text = TxtTransSerial.text
        .Text2.text = XPTxtBillID.text
    
        For RowNum = 1 To Fg.Rows - 1

            If .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Code")) <> "" Then
                .Fg.Rows = .Fg.Rows + 1
            End If

            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Name")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Name")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Name")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Code")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("ItemCase")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("HaveSerial")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("HaveSerial")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("HaveSerial")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Count")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("Price")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("Price")))
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("DiscountType")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 20) AND (dbo.TblItemsUnits.ItemID = " & Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .Fg.Cell(flexcpData, RowNum, .Fg.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .Fg.TextMatrix(RowNum, .Fg.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))

            '        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
            '        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
            '           StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            '        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = 1 'FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
            '        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = "Ã—«„" 'FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))

        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdInfo_Click()
    Me.PopupMenu mdifrmmain.MnuInvPurchase
End Sub

Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    If XPTxtValue(1).text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.text = "R" Then
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
        Else
            .Tag = "N"
            .Txt(1).text = XPTxtValue(1).text
            .LblNoteID.Caption = XPTxtSerial(1).text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).text = val(LblPrecenValue.Caption)
            .Txt(5).text = val(LblInstallCount.Caption)

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            .Txt(7).text = val(LblInstallSeprator.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                .OptInt(2).value = True
            End If

            With .Fg
                .Rows = Me.FgInstallments.Rows

                For i = 1 To Me.FgInstallments.Rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Value")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Value"))
                    .TextMatrix(i, .ColIndex("Due_Date")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Due_Date"))
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdNotes_Click()
    ShowRelatedNotes val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Dim StrTemp As String

    If val(Me.CmdNotes.Tag) = 0 Then
        Me.CmdNotes.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & val(Me.CmdNotes.Tag)
        Me.CmdNotes.ToolTipText = StrTemp
    End If

End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim StrTemp As String

    If val(Me.CmdRetruns.Tag) = 0 Then
        Me.CmdRetruns.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & val(Me.CmdRetruns.Tag)
        Me.CmdRetruns.ToolTipText = StrTemp
    End If

End Sub

Private Sub Cmmadd_Click()
    'Dim D As Double
    'D = Me.Grid.TextMatrix(1, Me.Grid.ColIndex("select"))
    'Dim I As Integer
    '
    'Txt_EXport.text = 0
    '     For I = 1 To Grid.Rows - 1
    '
    '        If Val(Grid.TextMatrix(I, Grid.ColIndex("select"))) = -1 Then
    '
    '        Txt_EXport.text = Val(Txt_EXport.text) + Val(Grid.TextMatrix(I, Grid.ColIndex("note_value")))
    '
    '        End If
    '        Next
End Sub

Private Sub DBCboClientName_Change()
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                CboPaymentType.locked = True
                '  CboPayMentType.ListIndex = 0
            Else
                CboPaymentType.locked = False
            End If
        End If
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        StrSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not (IsNull(RsTemp("Trans_DiscountTypePur").value)) Then
                If RsTemp("Trans_DiscountTypePur").value = 0 Then
                    Me.XPCboDiscountType.ListIndex = 0
                    Me.XPTxtDiscountVal.text = 0
                ElseIf RsTemp("Trans_DiscountTypePur").value = 1 Then
                    Me.XPCboDiscountType.ListIndex = 1
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                ElseIf RsTemp("Trans_DiscountTypePur").value = 2 Then
                    Me.XPCboDiscountType.ListIndex = 2
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                End If

            Else
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If

        Else
            Me.XPCboDiscountType.ListIndex = 0
            Me.XPTxtDiscountVal.text = 0
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)
        
    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 4
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 6

            If Me.WindowState = vbNormal Then
                Me.WindowState = vbMaximized
            Else
                Me.WindowState = vbNormal
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()
    Set m_MnuShowNewItemsPrices = mdifrmmain.MnuInvPurchaseMnu2
    Set m_MenuViewList = mdifrmmain.MnuInvPurchaseMnu1
    Set m_MenuShowItemCostEffect = mdifrmmain.MnuInvPurchaseMnu4
End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.text), 1
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    'With Me.Grid
    '    Select Case FG.ColKey(Col)
    '        Case "Chose"
    '            Cancel = False
    '        End Select
    'End With
End Sub

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_MenuShowItemCostEffect_Click()

    If Me.TxtModFlg.text = "R" Then
        ShowItemCostEffectForTrans 1, , Trim$(Me.TxtTransSerial.text)
    End If

End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim Fg As VSFlex8UCtl.vsFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set Fg = FrmView.vsfGroup1.vsFlexGrid

    With Fg
        .Cols = 9
        .RowHeightMin = 320
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
    
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"

        ',
        'QryTransactionsTotal.TransSum
        'QryTransactionsTotal.TransNet,
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT TOP 100 PERCENT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.Transaction_Date, " & "dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, dbo.TblStore.StoreName," & "QryTransactionsTotal.Trans_DiscountType,QryTransactionsTotal.Trans_Discount ," & "QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN "
            StrSQL = StrSQL + "dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID " & "LEFT OUTER JOIN dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = 1)"
            StrSQL = StrSQL + " ORDER BY QryTransactionsTotal.Transaction_ID "
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type= 1 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        If BolFrmLoaded = True Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"
        .ColKey(8) = "TotalAfterTax"
        'Rs.Close
        'Set Rs = Nothing
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
    FrmView.vsfGroup1.update
    FrmView.BolRetrunOnDblClick = True
    FrmView.SetDblClickRetrun Me, "Transaction_ID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·›Ê« Ì— «·„‘ —Ì« "
    FrmView.show
End Sub

Private Sub m_MnuShowNewItemsPrices_Click()

    If Not NewGrid Is Nothing Then
        NewGrid.ShowNewItemsPrice
    End If

End Sub

Private Sub Txt_EXport_GotFocus()
    'On Error GoTo ErrTrap

    'With Me.Grid
    '    .Rows = .FixedRows
    '    .ExtendLastCol = True
    '    .RowHeightMin = 300
    '    .Editable = flexEDKbdMouse
    '    .ExplorerBar = flexExSortShowAndMove
    '
    '    .AutoSize 0, .Cols - 1, False
    'End With

    'Dim I As Integer
    'Dim Rs As ADODB.Recordset
    'Dim My_SQL As String
    '
    'Set Rs = New ADODB.Recordset
    '
    'My_SQL = "SELECT dbo.Notes.NoteID , dbo.Notes.Note_Value, dbo.ExpensesType.Name FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3)"
    '
    ''    My_SQL = "select * From TblEmployee  where DateEndPasp < getdate()"
    '
    'Rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    'With Me.Grid
    '    .Rows = 2
    '    .Clear flexClearScrollable
    '    If Rs.RecordCount > 0 Then
    '        .Rows = Rs.RecordCount + 1
    '        Rs.MoveFirst
    '        For I = 1 To .Rows - 1
    ''             .TextMatrix(i, .ColIndex("Ser")) = i
    ''
    '             .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs.Fields("Name").Value), _
    '            "", Rs.Fields("Name").Value)
    '
    '            .TextMatrix(I, .ColIndex("NoteID")) = IIf(IsNull(Rs.Fields("NoteID").Value), _
    '            "", Rs.Fields("NoteID").Value)
    '
    '                        .TextMatrix(I, .ColIndex("Note_Value")) = IIf(IsNull(Rs.Fields("Note_Value").Value), _
    '            "", Rs.Fields("Note_Value").Value)
    '
    '            Rs.MoveNext
    '        Next
    '       Rs.Close
    '    End If
    '    .RowHeight(-1) = 300
    'End With
    'ErrTrap:

    'Dim StrSQL As String
    'Dim i As Double
    '
    'StrSQL = "SELECT dbo.Notes.NoteID , dbo.Notes.Note_Value, dbo.ExpensesType.Name FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3)"
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '
    '
    '
    'Rs.MoveFirst
    '   For i = 0 To Rs.RecordCount - 1
    '
    '   lstExp.AddItem Rs("NoteID") & Space$(5) & Rs("Note_Value") & Space$(5) & Rs("Name")
    '
    '    Rs.MoveNext
    '
    '    Next
    'End If

End Sub

Private Sub Txt_EXport_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Txt_EXport.ToolTipText = "„Ã„Ê⁄ «·„’—Ê›«  « Ê„« ÌﬂÌ« ⁄·Ï «–‰ «·«÷«›… "
End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPFillData_Click
            End If
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                Fg.SetFocus
            End If
        End If
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

Private Sub ChangeLang()

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    'Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.Caption = "Recive Voucher"
    Ele(6).Caption = Me.Caption
    lbl(8).Caption = "Invoice ID"
    lbl(7).Caption = "Invoice Date"
    lbl(6).Caption = "Vendor Name"
    lbl(4).Caption = "Store "

    'lbl(25).Caption = "Employee "
    lbl(10).Caption = "Payment Type"
    lbl(5).Caption = "Discount Type"
    lbl(11).Caption = "Value"

    Label1.Caption = "Another Expenses"
    CmdConvert.Caption = "Convert to bill"

    'lbl(22).Caption = "Profit Value"
    'lbl(23).Caption = "Profit Perce"

    lbl(3).Caption = " Total:"
    lbl(50).Caption = "Disc"
    lbl(24).Caption = " Net:"

    lbl(1).Caption = " By:"
    lbl(0).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"
    lbl(32).Caption = "Sales Type"
    lbl(33).Caption = "Cash Customer"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    Me.XPTab301.TabCaption(0) = "Items"
    
    Me.XPTab301.TabCaption(1) = "Notes"
    lbl(20).Caption = "Payment Method"
    XPChkPayType(0).Caption = "Cahs"
    XPChkPayType(1).Caption = "Due"
    XPChkPayType(0).Caption = "Check"
    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(12).Caption = "Serial"
    lbl(14).Caption = "Serial"
    '    lbl(11).Caption = "Box Name"
    lbl(21).Caption = "Due Date"
    
    lbl(18).Caption = "Check NO."
    lbl(17).Caption = "Bank Name"
    lbl(19).Caption = "Due Date"
    CmdINSTALLMENT.Caption = "INSTALLMENT"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.Ele(15).Caption = "Write any Comments about this Invoice"

    With Me.Fg
        .TextMatrix(0, .ColIndex("NewItem")) = "NewItem"
    End With
 
End Sub

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim Dcombos As ClsDataCombos
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    SetDtpickerDate XPDtbBill
    Set NewGrid = New ClsGrid
    NewGrid.GridTrans = PurchaseTransaction
    Set NewGrid.Grid = Me.Fg
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = Me.TxtModFlg
    Set NewGrid.TxtTotal = Me.XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.STORENAME = Me.DCboStoreName
    Set NewGrid.GrdTBar = Me.TBar
    '-----------------------------------------------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    '-----------------------------------------------------------------------------
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DcboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = cmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.LblTaxSalesValue = Me.lbl(25)
    Set NewGrid.LblTaxAddValue = Me.lbl(32)
    Set NewGrid.LblTaxStampValue = Me.lbl(33)
    Set NewGrid.LblTaxServiceValue = Me.lbl(49)

    Fg.WallPaper = BGround.Picture

    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    If SystemOptions.UserInterface = EnglishInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem "NO"
            .AddItem "Value  "
            .AddItem "Percentage"
        End With

        With CboPaymentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With

    Else

        With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        End With

        With CboPaymentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

    End If

    NewGrid.FillGrid
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID
    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
    cSearchDcbo(2).SetBuddyText Me.TxtStoreID
    '-----------------------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    Me.XPChkTAX.value = vbUnchecked
    XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    Me.ChkTaxStamp.value = vbUnchecked
    ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click
    '-----------------------------------------------------------------------------
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type= 20 and Transaction_Type_Sub=28"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    Retrive
    Me.TxtModFlg.text = "R"
    Resize_Form Me, TransactionSize

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set BuyReport = Nothing
    Set m_MnuShowNewItemsPrices = Nothing

    If Not m_FrmSearch Is Nothing Then
        Unload m_FrmSearch
        Set m_FrmSearch = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            Me.Caption = "«–‰ «÷«›… «’‰«› „’‰⁄Â ··«‰ «Ã «· «„"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
     
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = False
        
            XPCboDiscountType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.XPTxtDiscountVal.locked = True
        
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
        
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            Fg.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If
        
            CboPaymentType.locked = True
            DtpDelayDate.Enabled = False
            Ele(4).Enabled = False
        
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False
        
        Case "N"
            Me.Caption = "«–‰ «÷«›… «’‰«› „’‰⁄Â ··«‰ «Ã «· «„ ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '    Me.XPBtnMove(0).Enabled = False
            '    Me.XPBtnMove(1).Enabled = False
            '    Me.XPBtnMove(2).Enabled = False
            '    Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            Fg.Enabled = True
            Fg.Rows = 2
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            Fg.Editable = flexEDKbdMouse
            XPDtbBill.value = Date
            '        XPFillData.Enabled = True
            XPCboDiscountType.ListIndex = 0
            CboPaymentType.ListIndex = 0
            CboPaymentType.locked = False
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            Ele(4).Enabled = True
        
            CboItemCase.ListIndex = 0
        
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True

        Case "E"
            Me.Caption = "«–‰ «÷«›… «’‰«› „’‰⁄Â ··«‰ «Ã «· «„ (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
        
            Fg.Enabled = True
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            DtpDelayDate.Enabled = True
        
            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPaymentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            Fg.Editable = flexEDKbdMouse
        
            CboPaymentType.locked = False
            DBCboClientName_Change
            Ele(4).Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
    End Select

    Exit Sub
ErrTrap:
    Stop
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest As ADODB.Recordset
    Dim Num As Long
    Dim Msg As String
    Dim i As Integer
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    'Dim rs As ADODB.Record
    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""
    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.text = ""

    '---------------------------------------------
    '---------------------------------------------
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.TxtNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)
    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", Trim(rs("Trans_Discount").value))
    CboPaymentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    '÷—»Ì… «·„»Ì⁄« 
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    Dim Myrec As New ADODB.Recordset
    Dim Mytotal As Integer
    Dim MySQL As String

    MySQL = "SELECT Sum (Notes.Note_Value) AS [TotalRevenue] FROM Notes where NumOrderInpot = " & TxtTransSerial
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.BOF Or RsNotes.EOF) Then
        Txt_EXport.text = IIf(IsNull(RsNotes("TotalRevenue").value), "", (RsNotes("TotalRevenue").value))
    End If

    Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    'Txt_EXport.text = IIf(IsNull(rs("Shahne").Value), "", (rs("Shahne").Value))

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = rs("TaxServiceValue").value
        End If
    End If

    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    XPTxtSum.text = ""

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + "  where Transaction_ID=" & val(rs("Transaction_ID").value)

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showQty")), "", (RsDetails("showQty").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                Fg.TextMatrix(Num, Fg.ColIndex("HaveSerial")) = True
            End If

            Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))

            If Fg.TextMatrix(Num, Fg.ColIndex("Price")) = "" Then
                Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            End If

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If

            Fg.TextMatrix(Num, Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            Fg.TextMatrix(Num, Fg.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            Fg.TextMatrix(Num, Fg.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").value))
        
            Fg.Cell(flexcpData, Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext

            If Fg.Rows > 10 Then
                If Num = 8 Then Fg.Refresh
            End If

        Next Num

        Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""

    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    DtpDelayDate.value = Date
    StrSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                'Me.TxtNoteID(0).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                'Me.TxtNoteID(1).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 13 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
            End If
        
            RsNotes.MoveNext
        Next Num

    End If

    Set RsNotes = New ADODB.Recordset
    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
    StrSQL = StrSQL + " Where NoteType=13 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + " Order BY Notes.NoteID"
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgCheques
        .Rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .Rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)

                If Not IsNull(RsNotes("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
                Else
                    .TextMatrix(i, .ColIndex("DueDate")) = ""
                End If

                RsNotes.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
        SumChecks
    End With

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            CmdINSTALLMENT.Enabled = True
            CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            LngPartID = RsTest("PartID").value
            Me.LblPrecenType.Tag = RsTest("InterestType").value

            If RsTest("InterestType").value = 0 Then
                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
            ElseIf RsTest("InterestType").value = 1 Then
                LblPrecenType.Caption = "ﬁÌ„… À«» …"
            ElseIf RsTest("InterestType").value = 2 Then
                LblPrecenType.Caption = "·«ÌÊÃœ"
            End If

            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
            Me.LblInstallTotal.Caption = RsTest("Total").value
            Me.LblInstallCount.Caption = RsTest("InstallCount").value
            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value

            If RsTest("InstallmentType").value = 0 Then
                LblInstallmentType.Caption = "ÌÊ„"
            ElseIf RsTest("InstallmentType").value = 1 Then
                LblInstallmentType.Caption = "‘Â—"
            ElseIf RsTest("InstallmentType").value = 2 Then
                LblInstallmentType.Caption = "”‰…"
            End If

            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
            Set RsPartDetails = New ADODB.Recordset
            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .Rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .Rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)

                        If Not IsNull(RsPartDetails("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
                        Else
                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
                        End If
 
                        RsPartDetails.MoveNext
                    Next i

                End With

            End If

        Else
            CmdINSTALLMENT.Enabled = False
            CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
        End If
    End If

    NewGrid.Calculate 1, , , True
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    TxtFillData.text = "F"
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Msg = "Œÿ« ›Ï ≈” —Ã«⁄ «·»Ì«‰« ..!!!"
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String
    Dim BegainTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtBillID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If AvailableDeal = True Then
                If Not rs.RecordCount < 1 Then
                    Cn.BeginTrans
                    BegainTrans = True
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                    Cn.Execute StrSQL, , adExecuteNoRecords
                
                    StrSQL = "delete From Notes where noteid=" & val(TxtNoteID.text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
         
                    rs.delete
                    Cn.CommitTrans
                    BegainTrans = False
                    rs.MoveFirst

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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Set TTP = New clstooltip
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… «–‰ «÷«›… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «–‰ «·«÷«›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «–‰ «·«÷«›…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «–‰ «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄„·Ì«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… «–‰ «·«÷«›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnAdd, _
    '    "≈÷«›… «·√’‰«› ..." & Wrap & _
    '    " ·«÷«›… ’‰› ÃœÌœ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F2", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnRemove, _
    '    "Õ–› ’‰› ..." & Wrap & _
    '    "·Õ–› √Õœ «·√’‰«›" & Wrap & _
    '    " ÕœœÂ Ê«÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F3", True
    'End With
    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPFillData, _
    '    " ⁄»∆… »Ì«‰«  «·√’‰«›" & Wrap & _
    '    "· ⁄»∆… »Ì«‰«  «·√’‰«› ›Ì" & Wrap & _
    '    "›Ì ‰«›–… ÕÊ«—" & Wrap & _
    '    "  ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— Ctrl + Space", True
    'End With
    With TTP
        .Create Me.hwnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Msg As String
    Dim Mytot As String
    Dim RowNum As Integer
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    Dim note_id As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String

    On Error GoTo ErrTrap

    If Trim(Me.TxtTransSerial.text) = "" Then
        Msg = "ÌÃ» ﬂ «»… —ﬁ„ «–‰ «·«÷«›… ..!!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtTransSerial.SetFocus
        Exit Sub
    End If

    If Me.TxtModFlg.text = "N" Then
        If RepeatSerial(Trim(Me.TxtTransSerial.text), 20, 0, val(Me.DBCboClientName.BoundText)) = True Then
            Exit Sub
        End If

    ElseIf Me.TxtModFlg.text = "E" Then

        If RepeatSerial(Trim(Me.TxtTransSerial.text), 1, val(Me.XPTxtBillID.text), val(Me.DBCboClientName.BoundText)) = True Then
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbArrowHourglass

    If DBCboClientName.text = "" Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If DCboStoreName.text = "" Then
        Msg = "„‰ ›÷·ﬂ Õœœ «”„ «·„Œ“‰"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            Msg = "ÌÃ»  ÕœÌœ ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountVal.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Not IsNumeric(XPTxtDiscountVal.text) Then
            Msg = "ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountVal.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        XPTxtDiscountVal.SetFocus
    End If

    If CboPaymentType.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPaymentType.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPChkPayType(0).value = vbChecked Then
        If Me.DcboBox.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.TxtModFlg.text = "N" Then
            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), Me.XPDtbBill.value) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), Me.XPDtbBill.value, , , val(Me.XPTxtValue(0).Tag)) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    If val(Me.XPTxtValue(1).text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If val(Me.LblInstallTotal.Caption) <> val(Me.XPTxtValue(1).text) Then
                Me.XPTxtValue(1).text = val(Me.LblInstallTotal.Caption)
            End If
        End If
    End If

    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If DBCboClientName.BoundText = 1 Then
    '    MsgBox "ÌÃ» «Œ Ì«— „Ê—œ √Œ—"
    ' Exit Sub
    'End If

    'Check the Items Grid
    XPTab301.CurrTab = 0

    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    If Me.TxtModFlg.text = "E" Then
    '        If EditTransStatus(Val(Me.XPTxtBillID.text), "E", NewGrid) = False Then
    '            Exit Sub
    '        End If
    '    End If
    '---------------------------------------------------------------
    Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS where Transaction_ID = " & val(Text2.text)

    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '-------------------------------
    If Me.XPChkPayType(0).value = vbChecked Then
        DblNotesTotal = val(Me.XPTxtValue(0).text)
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).text)
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.lbl(18).Caption)
    End If

    'If DblNotesTotal <> Val(LblTotal.Caption) Then
    '    Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ ”«ÊÏ „⁄ ≈Ã„«·Ï «·›« Ê—…...!!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If

    Set RsNotesGeneral = New ADODB.Recordset
    RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    If Me.TxtModFlg.text = "N" Then
    
    Else
        '   StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
        '   Cn.Execute StrSqlDel, , adExecuteNoRecords
        '   StrSqlDel = "delete From Notes where Transaction_ID=" & Val(rs("Transaction_ID").value)
        '   Cn.Execute StrSqlDel, , adExecuteNoRecords
        '
        '   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & Val(Me.XPTxtBillID.text)
        '   Cn.Execute StrSQL, , adExecuteNoRecords
        '
        StrSqlDel = "delete From Notes where noteid=" & val(TxtNoteID.text)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        general_noteid = val(TxtNoteID.text)
    End If

    RsNotesGeneral.AddNew
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    general_noteid = RsNotesGeneral("NoteID").value
    TxtNoteID.text = general_noteid
    ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
    RsNotesGeneral("NoteDate").value = XPDtbBill.value
    RsNotesGeneral("NoteType").value = 160 ' «–‰ «÷«›…
    RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
    RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    RsNotesGeneral("numbering_type1").value = sand_numbering_type(9) '  «–‰ «÷«›…
    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    RsNotesGeneral.update
        
    '---------Start Saving------------------------------------------------
    Set RSTransDetails = New ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
    RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    BeginTrans = True

    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
        rs("Transaction_ID").value = val(XPTxtBillID.text)
    ElseIf Me.TxtModFlg.text = "E" Then

        If rs("Transaction_ID").value <> val(XPTxtBillID.text) Then
            rs.find "Transaction_ID=" & val(XPTxtBillID.text), , adSearchForward, 1
        End If

        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("NoteId").value = val(TxtNoteID.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", Null, Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Type").value = 20 '1
    rs("Transaction_Type_Sub").value = 28

    rs("UserID").value = user_id
    rs("Shahne").value = val(Txt_EXport.text)
    rs("nots").value = Text1.text

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

    If XPCboDiscountType.ListIndex = -1 Or XPCboDiscountType.ListIndex = 0 Then
        rs("Trans_Discount").value = Null
    Else
        rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
    End If

    If CboPaymentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPaymentType.ListIndex)
    End If

    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, (DCboStoreName.BoundText))
    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    rs.update

    If Me.TxtModFlg.text = "E" Then
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
    End If

    For RowNum = 1 To Fg.Rows - 1

        'Check Repeat Serial
        If Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")) <> "" Then
            StrSQL = "select * From Transaction_Details where ItemSerial='" & Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")) & "'"
            StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                Msg = Msg + Fg.Cell(flexcpTextDisplay, RowNum, Fg.ColIndex("name")) & Chr(13)
                Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                RsTemp.Close
                XPTab301.CurrTab = 0
                Fg.Row = RowNum
                Fg.Col = Fg.ColIndex("name")
                Fg.ShowCell RowNum, Fg.ColIndex("name")
                Fg.SetFocus
                Screen.MousePointer = vbDefault
                BeginTrans = False
                Cn.RollbackTrans
                Exit Sub
            End If

            RsTemp.Close
        End If

        If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
            RSTransDetails.AddNew
            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
            RSTransDetails("Item_ID").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code"))))
            RSTransDetails("Quantity").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))))

            '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) _
            '            = ""), "", Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
            If Not Fg.TextMatrix(RowNum, Fg.ColIndex("Name")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If RsTemp("HaveSerial").value = True Then
                        RSTransDetails("ItemSerial").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")) = ""), Null, (Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))))
                    End If
                End If

                RsTemp.Close
            End If

            RSTransDetails("ItemCase").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")) = ""), Null, (Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase"))))
            RSTransDetails("showPrice").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")) = ""), Null, (Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType"))))
            RSTransDetails("ItemDiscount").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal"))))
            RSTransDetails("guaranteeTime").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("guaranteeTime")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("guaranteeTime"))))
        
            RSTransDetails("Remarks").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Remarks")) = ""), Null, Trim$(Fg.TextMatrix(RowNum, Fg.ColIndex("Remarks"))))
         
            '.TextMatrix(LngRow, .ColIndex("ColorID")) = 1
            '.TextMatrix(LngRow, .ColIndex("ItemSize")) = 0
        
            RSTransDetails("ColorID").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ColorID")) = ""), 1, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ItemSize")) = ""), "", Trim$(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemSize"))))
            
            RSTransDetails("UnitID").value = IIf(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")) = "", Null, (Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID"))))
            RSTransDetails("ShowQty").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))))
             
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
            LngUnitID = val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")))
            DblQty = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            End If
        
            '        Dim RsPrice As New ADODB.Recordset
            '        Set RsPrice = New ADODB.Recordset
            '
            '        RsPrice.Open "select UnitPurPrice from TblItemsUnits where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & " and UnitID=" & FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")), Cn, adOpenStatic, adLockOptimistic, adCmdTable
            
            ' RSTransDetails("price").Value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").Value, 2)
            If val(Txt_EXport.text) > 0 Then
                Dim mm As String
                Dim Myprc As String
                mm = MsgBox(" Â· Â‰«ﬂ „’«—Ì› √Œ—Ï ⁄·Ï Â–« «·«–‰ ... «–«  „  Õ„Ì· Â–Â «·„’—Ê›«  ›·« ÌÕﬁ ·ﬂ «· ⁄œÌ·", vbYesNo)

                If mm = vbYes Then

                    '   Â·  „  ÕÊÌ· «·«–‰ «·Ï ›« Ê—…
                    If Text1.text <> "" Then
                            
                        RSTransDetails("ToTAlELSHahn") = (((RSTransDetails("showPrice") * RSTransDetails("ShowQty")) / val(LblTotal.Caption)) * val(Txt_EXport.text)) / RSTransDetails("ShowQty")
                      
                        Myprc = RSTransDetails("showprice").value / RSTransDetails("QtyBySmalltUnit").value
                         
                        Myprc = (RSTransDetails("ToTAlELSHahn").value / RSTransDetails("QtyBySmalltUnit").value) + Myprc
                        RSTransDetails("Price").value = Myprc
                               
                        Mytot = RSTransDetails("showprice").value + RSTransDetails("ToTAlELSHahn")
                        RSTransDetails("showprice").value = Mytot
                    Else
                        MsgBox "ÌÃ»  ÕÊÌ· «·«–‰ «·Ï ›« Ê—… ﬁ»·  Õ„Ì·Â »«· ﬂ«·Ì› «·«Œ—Ï"
                    End If
                End If

                ' Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").Value, 2)
            End If

            RSTransDetails.update
        End If

    Next RowNum

    '------------------------------------------------------------------------------
    '„‰ Â‰« «·ﬂÊœ  Ê«ﬁ›
    '------------------------------------------------------------------------------
    'If Me.XPChkPayType(0).Value = Checked Then
    '    RsNotes.AddNew
    '    RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '    Note_ID = RsNotes("NoteID").Value
    '    If Me.TxtModFlg.text = "N" Then
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    '    ElseIf Trim(XPTxtSerial(0).text) <> "" Then
    '        RsNotes("NoteSerial").Value = Trim(XPTxtSerial(0).text)
    '    Else
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    '    End If
    '    RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '    RsNotes("NoteDate").Value = XPDtbBill.Value
    '    RsNotes("NoteType").Value = 0
    '    RsNotes("Note_Value").Value = _
    '    IIf(XPTxtValue(0).text = "", Null, Val(XPTxtValue(0).text))
    '    RsNotes("Member_ID").Value = _
    '    IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '    RsNotes("BankID").Value = Null
    '    RsNotes("BoxID").Value = IIf(DcboBox.BoundText = "", Null, Val(DcboBox.BoundText))
    '    RsNotes("CusID").Value = Null
    '    RsNotes.update
    '    Me.XPTxtValue(0).Tag = RsNotes("NoteID").Value
    '    '--------------------------------------------------------------------------
    'End If
    'If Me.XPChkPayType(1).Value = Checked Then
    RsNotes.AddNew
    RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    note_id = RsNotes("NoteID").value
    RsNotes("NoteDate").value = XPDtbBill.value

    If Me.TxtModFlg.text = "N" Then
        RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
        XPTxtSerial(1).text = RsNotes("NoteSerial").value
    ElseIf Trim(XPTxtSerial(1).text) <> "" Then
        RsNotes("NoteSerial").value = Trim(XPTxtSerial(1).text)
    Else
        RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
        XPTxtSerial(1).text = RsNotes("NoteSerial").value
    End If

    RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
    RsNotes("NoteType").value = 1
    RsNotes("Note_Value").value = val(LblTotalAll.Caption)
    
    RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    RsNotes("BankID").value = Null
    RsNotes("CusID").value = Null
    RsNotes("DueDate").value = DtpDelayDate.value
    RsNotes.update
    Me.XPTxtValue(1).Tag = RsNotes("NoteID").value
    'End If
    'If Me.XPChkPayType(2).Value = Checked Then
    '    With Me.FgCheques
    '        For I = .FixedRows To .Rows - 1
    '            RsNotes.AddNew
    '                RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '                Note_ID = RsNotes("NoteID").Value
    '                RsNotes("NoteDate").Value = XPDtbBill.Value
    '                RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '                RsNotes("NoteType").Value = 13
    '                RsNotes("Note_Value").Value = Val(.TextMatrix(I, .ColIndex("CheckValue")))
    '                RsNotes("BankID").Value = Val(.TextMatrix(I, .ColIndex("BankID")))
    '                RsNotes("ChqueNum").Value = Trim$(.TextMatrix(I, .ColIndex("CheckNumber")))
    '                RsNotes("DueDate").Value = CDate(Trim$(.TextMatrix(I, .ColIndex("DueDate"))))
    '                RsNotes("Member_ID").Value = Val(Me.DBCboClientName.BoundText)
    '                RsNotes("CUSID").Value = Val(Me.DBCboClientName.BoundText)
    '            RsNotes.update
    '            '--------------------------------------------------------------------------
    '        Next I
    '    End With
    'End If
    'Õ›Ÿ «·√›”«ÿ
    'If Me.XPChkPayType(1).Value = Checked Then
    '    If ChkInstall.Value = vbChecked Then
    '        'Save installment Data
    '        Set RsTemp = New ADODB.Recordset
    '        RsTemp.Open "InstallMent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '        RsTemp.AddNew
    '            RsTemp("PartID").Value = CStr(new_id("InstallMent", "PartID", "", True))
    '            RsTemp("NoteID").Value = Note_ID
    '            RsTemp("BasicAmmount").Value = IIf(XPTxtValue(1).text = "", 0, Val(XPTxtValue(1).text))
    '            RsTemp("InterestType").Value = Val(Me.LblPrecenType.Tag)
    '            RsTemp("InterestVal").Value = Val(LblPrecenValue.Caption)
    '            RsTemp("Total").Value = Val(LblInstallTotal.Caption)
    '            RsTemp("InstallCount").Value = Val(LblInstallCount.Caption)
    '            RsTemp("FirstInstallDate").Value = CDate(Me.LblFirstInstallDate.Caption)
    '            If Val(LblInstallmentType.Tag) = 0 Then
    '                RsTemp("InstallmentType").Value = 0
    '            ElseIf Val(LblInstallmentType.Tag) = 1 Then
    '                RsTemp("InstallmentType").Value = 1
    '            ElseIf Val(LblInstallmentType.Tag) = 2 Then
    '                RsTemp("InstallmentType").Value = 2
    '            End If
    '            RsTemp("InstallSeprator").Value = Val(Me.LblInstallSeprator.Caption)
    '            RsTemp("StartValue").Value = IIf(Val(Me.LblStartValue.Caption) = 0, Null, Val(Me.LblStartValue.Caption))
    '            RsTemp("CustID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '            RsTemp("Type").Value = 1
    '        RsTemp.update
    '        'save installment Details
    '        Set RsDetalis = New ADODB.Recordset
    '        RsDetalis.Open "InstallMentDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '        With Me.FgInstallments
    '            For RowNum = 1 To .Rows - 1
    '                RsDetalis.AddNew
    '                    RsDetalis("QestID").Value = CStr(new_id("InstallMentDetails", "QestID", "", True))
    '                    RsDetalis("PartID").Value = RsTemp("PartID").Value
    '                    RsDetalis("QeqtNum").Value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", "", .TextMatrix(RowNum, .ColIndex("Serial")))
    '                    RsDetalis("Value").Value = IIf(.TextMatrix(RowNum, .ColIndex("Value")) = "", "", Val(.TextMatrix(RowNum, .ColIndex("Value"))))
    '                    RsDetalis("DueDate").Value = IIf(.TextMatrix(RowNum, .ColIndex("Due_Date")) = "", "", .TextMatrix(RowNum, .ColIndex("Due_Date")))
    '                    RsDetalis("Receipt").Value = False
    '                RsDetalis.update
    '            Next RowNum
    '        End With
    '    End If
    'End If
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Single
    Dim Account_Code_dynamic As String

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = " ”‰œ «” ·«„ «’‰«›  „‰ Ã  «„ —ﬁ„ " & Me.TxtTransSerial.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ «” ·«„ «’‰«›  „‰ Ã  «„  —ﬁ„ " & Me.TxtTransSerial.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With Fg

                For i = 1 To Fg.Rows - 1

                    If Fg.TextMatrix(i, Fg.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(Fg.TextMatrix(i, Fg.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = Fg.TextMatrix(i, Fg.ColIndex("Price")) * Fg.TextMatrix(i, Fg.ColIndex("Count"))
    
                        StrTempDes = "«”‰œ «” ·«„ «’‰«›  „‰ Ã  «„  —ﬁ„ " & Me.TxtTransSerial.text
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(4, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„‘ —Ì«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
            
                StrTempDes = "«–‰ «÷«›…" & Me.TxtTransSerial.text
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With Fg

                    For i = 1 To Fg.Rows - 1

                        If Fg.TextMatrix(i, Fg.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(Fg.TextMatrix(i, Fg.ColIndex("Code")), val(my_branch), 4)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   «·„‘ —Ì«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = Fg.TextMatrix(i, Fg.ColIndex("Price")) * Fg.TextMatrix(i, Fg.ColIndex("Count"))
    
                            StrTempDes = "«–‰ «÷«›…" & Me.TxtTransSerial.text
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If

        '
        '        Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
        '        If Account_Code_dynamic = "" Then
        '         MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
        '        GoTo ErrTrap
        '        End If
        '
        '    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…0
        '  '  StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
        '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
        '    LngDevNO = LngDevNO + 1
        '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '        0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '        GoTo ErrTrap
        '    End If
    End If

    If XPChkTAX.value = vbChecked Then
        '  StrTempAccountCode = "a1a3a5" '÷—»Ì… „»Ì⁄«  „œÌ‰…
        '  SngTemp = Val(Me.lbl(25).Caption)
        '  StrTempDes = "«–‰ «÷«›…  —ﬁ„ " & Me.TxtTransSerial.text
        '  LngDevNO = LngDevNO + 1
        '  If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '     0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '      GoTo ErrTrap
        '  End If
    End If

    If Me.ChkTaxAdd.value = vbChecked Then
        '  StrTempAccountCode = "a2a5a4" '÷—»Ì… √—»«Õ  Ã«—Ì… (Œ’„ Ê≈÷«›…
        '  StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
        '  SngTemp = Val(Me.lbl(32).Caption)
        '  LngDevNO = LngDevNO + 1
        '  If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '      0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '      GoTo ErrTrap
        '  End If
    End If

    '«·œ«∆‰
    'If CboPaymentType.ListIndex = 0 Then  'Me.XPChkPayType(0).Value = vbChecked Then
    '    '«·Œ“Ì‰…
    '    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
    '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '
    '    SngTemp = DisplayCurrency(Val(Me.XPTxtValue(0).text))
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    'If CboPaymentType.ListIndex = 1 Then 'Me.XPChkPayType(1).Value = vbChecked Then
    '    '«·√Ã·
    '    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", Val(Me.DBCboClientName.BoundText))
    '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.lbltotal.Caption), _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    'If Me.XPChkPayType(2).value = vbChecked Then
    '  '  StrTempAccountCode = "a2a3a2" '√Ê—«ﬁ «·œ›⁄
    '  '  StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & Chr(13)
    '  '  StrTempDes = StrTempDes & "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '  '  LngDevNO = LngDevNO + 1
    '  '  If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.lbl(18).Caption), _
    '  '      1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '  '      GoTo ErrTrap
    '  '  End If
    'End If
    'If Val(Me.LblDiscountsTotal.Caption) > 0 Then
    '         Account_Code_dynamic = get_account_code_branch(13, my_branch)
    '
    '        If Account_Code_dynamic = "NO branch" Then
    '        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
    '        GoTo ErrTrap
    '        Else
    '        If Account_Code_dynamic = "NO account" Then
    '           MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·Œ’„ «·„ﬂ ”» ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
    '        GoTo ErrTrap
    '
    '        End If
    '        End If
    '    StrTempAccountCode = Account_Code_dynamic '«·Œ’„ «·„ﬂ ”»13
    '  '  StrTempAccountCode = "a4a4" '«·Œ’„ «·„ﬂ ”»
    '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblDiscountsTotal.Caption), _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If

    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & Chr(13)
    
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                Msg = " chages Was Saved " & Chr(13)
    
            End If

    End Select

    TxtModFlg.text = "R"
    UpdateTransCost val(Me.XPTxtBillID.text)

    If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
        '›Ï Õ«·… «‰  ﬂÊ‰ ÿ—Ìﬁ… Õ”«» „ Ê”ÿ «· ﬂ·›…
        'ÂÊ
        'ModernWeightAverage
        '·«»œ «‰ ÌﬁÊ„ «·»—‰«„Ã » ⁄œÌ· ﬁÌ„… „ Ê”ÿ «· ﬂ·›… ··√’‰«›
        '«·„ÊÃÊœ… ›Ï «·›« Ê—…
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    'Stop
    'Resume
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Msg = Msg & Err.description & Chr(13)
    Msg = Msg & Err.Number & Chr(13)
    Msg = Msg & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPBtnNewClients_Click()

    With FrmAddNewCustemer
        '    .Tag = "x"
        .DealingForm = PurchaseTransaction
        Set .DcboCustomers = DBCboClientName
        .Caption = "≈÷«›… „Ê—œ ÃœÌœ"
        .lbl(1).Caption = "ﬂÊœ «·„Ê—œ"
        .lbl(0).Caption = "«”„ «·„Ê—œ"
        .AddType = 2
        .show vbModal
    End With

End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
        lbl(11).Enabled = False
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.text = ""
    Else
        lbl(11).Enabled = True
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Fg.TextMatrix(1, Fg.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    DtpDelayDate.value = Date
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.Rows = Me.FgCheques.FixedRows
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        lbl(22).Enabled = True
        lbl(45).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(22).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub printing()
    On Error GoTo ErrTrap
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowBuyData XPTxtBillID.text
        End If

    Else

        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowBuyDataShort XPTxtBillID.text
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To Fg.Rows - 1

        If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & Fg.TextMatrix(RowNum, Fg.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If Fg.Cell(flexcpChecked, RowNum, Fg.ColIndex("HaveSerial")) = flexChecked Then
                If Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If Fg.Cell(flexcpChecked, RowNum, Fg.ColIndex("HaveSerial")) = flexChecked Then

                    With FrmAlarm
                        .Tag = "x"
                        .DealingForm = PurchaseTransaction
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    Set RsTemp = New ADODB.Recordset
                    LngItemID = val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("QTY").value) < val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = PurchaseTransaction
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "" Then Exit Sub
    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboPaymentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = XPTxtSum.text
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = ""
        End If
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPaymentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = XPTxtSum.text
    Exit Sub
ErrTrap:
End Sub

Public Function RepeatSerial(StrSerial As String, _
                             IntTransType As Integer, _
                             Optional IntTransID As Long = 0, _
                             Optional LngCusID As Long = 0) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    RepeatSerial = False

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Serial, " & "QryTransactionsTotal.Transaction_Date , QryTransactionsTotal.Transaction_Type," & "dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN " & "dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where QryTransactionsTotal.Transaction_Serial ='" & StrSerial & "'"
        StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Type=" & IntTransType & ""

        If LngCusID <> 0 Then
            StrSQL = StrSQL + " AND dbo.TblCustemers.CusID=" & LngCusID & ""
        End If

        If IntTransID <> 0 Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_ID <> " & IntTransID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Msg = "—ﬁ„ «·›« Ê—… „ÊÃÊœ „”»ﬁ« ›Ï «·»—‰«„Ã øø" & Chr(13)
            Msg = Msg + "„⁄·Ê„«  ⁄‰ «·›« Ê—… «·„”Ã·…:-" & Chr(13)
        
            Msg = Msg + "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã:" & rs("Transaction_ID").value & Chr(13)
            Msg = Msg + "„”·”· «·›« Ê—…:" & rs("Transaction_Serial").value & Chr(13)
            Msg = Msg + " «—ÌŒ  ”ÃÌ· «·›« Ê—…:" & rs("Transaction_Date").value & Chr(13)
            Msg = Msg + "«”„ «·⁄„Ì· «Ê «·„Ê—œ:" & rs("CusName").value & Chr(13)
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            RepeatSerial = True
        End If

        rs.Close
        Set rs = Nothing

    End If

End Function

Private Sub SetDefaults()
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
        XPDtbBill.value = Date
    ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
                XPDtbBill.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
            End If

            'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast

        If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
            XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), Date, (rs("Transaction_Date").value))
        End If
    End If

    Me.DcboBox.BoundText = 1
    Me.CboPaymentType.ListIndex = 1

End Sub
